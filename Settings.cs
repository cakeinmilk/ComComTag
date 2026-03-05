using System;
using System.IO;
using System.Collections.Generic;

namespace ComComTag {
    public class Settings {
        public string DefaultDirectory { get; set; }
        public List<string> Locations { get; set; }
        public string FFmpegPath { get; set; }
        public string DefaultBitrate { get; set; }

        private string _settingsFile;

        public Settings() {
            string appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            string appFolder = Path.Combine(appData, "ComComTag");
            if (!Directory.Exists(appFolder)) {
                Directory.CreateDirectory(appFolder);
            }
            _settingsFile = Path.Combine(appFolder, "settings.ini");

            // One-time migration from local directory to AppData
            string localSettings = "settings.ini";
            if (File.Exists(localSettings)) {
                try {
                    if (File.Exists(_settingsFile)) File.Delete(_settingsFile);
                    File.Move(localSettings, _settingsFile);
                } catch { } // Ignore if we can't move it for some reason
            }
            DefaultDirectory = "";
            Locations = new List<string> {
                "Bill Murray, Islington",
                "Soho Theatre, Soho",
                "Soho Theatre, Walthamstow",
                "Pleasance, Islington",
                "Up the Creek, Greenwich"
            };
            FFmpegPath = "ffmpeg.exe";
            DefaultBitrate = "128k";
            Load();
        }

        public void Load() {
            if (!File.Exists(_settingsFile)) {
                Save(); // Create default
                return;
            }

            var lines = File.ReadAllLines(_settingsFile);
            Locations.Clear();
            foreach (var line in lines) {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("DefaultDirectory=")) {
                    DefaultDirectory = trimmed.Substring("DefaultDirectory=".Length);
                } else if (trimmed.StartsWith("Location=")) {
                    Locations.Add(trimmed.Substring("Location=".Length));
                } else if (trimmed.StartsWith("FFmpegPath=")) {
                    FFmpegPath = trimmed.Substring("FFmpegPath=".Length);
                } else if (trimmed.StartsWith("DefaultBitrate=")) {
                    DefaultBitrate = trimmed.Substring("DefaultBitrate=".Length);
                }
            }
            if (Locations.Count == 0) Locations.Add("Default Location");
        }

        public void Save() {
            using (var writer = new StreamWriter(_settingsFile)) {
                writer.WriteLine(string.Format("DefaultDirectory={0}", DefaultDirectory));
                writer.WriteLine(string.Format("FFmpegPath={0}", FFmpegPath));
                writer.WriteLine(string.Format("DefaultBitrate={0}", DefaultBitrate));
                foreach (var loc in Locations) {
                    writer.WriteLine(string.Format("Location={0}", loc));
                }
            }
        }
    }
}
