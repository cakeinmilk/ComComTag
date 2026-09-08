using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace ComComTag {
    public class Settings {
        public string DefaultDirectory { get; set; }
        public List<string> Artists { get; set; }
        public List<string> Locations { get; set; }
        public string FFmpegPath { get; set; }
        public string DefaultBitrate { get; set; }
        public string InstaTempDirectory { get; set; }
        public bool FlushTempOnExit { get; set; }
        public bool AutoAddArtists { get; set; }
        public bool AutoAddLocations { get; set; }
        public bool UseFirstArtistAsAlbumArtist { get; set; }
        public bool GroupAdditionalArtists { get; set; }
        public int GroupAdditionalArtistsThreshold { get; set; }
        public string GroupAdditionalArtistsText { get; set; }
        public string DefaultM4bArtist { get; set; }
        public bool AppendChapterDuration { get; set; }
        public bool ClearErasesAllTags { get; set; }
        public string Theme { get; set; }
        public string DateFormat { get; set; }
        public string FilenamePattern { get; set; }

        // Conventions Playground persistence
        public string PlaygroundDate { get; set; }
        public string PlaygroundArtist { get; set; }
        public string PlaygroundShow { get; set; }
        public string PlaygroundLocation { get; set; }

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
                } catch { } // Ignore if we can't move it
            }

            DefaultDirectory = "";
            Artists = new List<string> {
                "Various Artists",
                "Comedian A",
                "Comedian B"
            };
            Locations = new List<string> {
                "Venue 1, City",
                "Venue 2, City",
                "Studio A",
                "Main Stage"
            };
            FFmpegPath = "ffmpeg.exe";
            DefaultBitrate = "128k";
            InstaTempDirectory = "";
            FlushTempOnExit = true;
            AutoAddArtists = true;
            AutoAddLocations = true;
            UseFirstArtistAsAlbumArtist = true;
            GroupAdditionalArtists = true;
            GroupAdditionalArtistsThreshold = 3;
            GroupAdditionalArtistsText = "and friends";
            DefaultM4bArtist = "Various Artists";
            AppendChapterDuration = true;
            ClearErasesAllTags = true;
            Theme = "System";
            DateFormat = "yyyy-MM-dd";
            FilenamePattern = "{Date} - {Artist} - {Show} - {Location}";

            // Dummy data defaults
            PlaygroundDate = "1999-06-13";
            PlaygroundArtist = "Wally Bazoom";
            PlaygroundShow = "Smile Time";
            PlaygroundLocation = "The Aigburth Arms, Liverpool";

            Load();
        }

        public string GetResolvedFFmpegPath() {
            if (string.IsNullOrWhiteSpace(FFmpegPath)) return "";
            if (Path.IsPathRooted(FFmpegPath)) return FFmpegPath;
            string appDir = AppDomain.CurrentDomain.BaseDirectory;
            return Path.Combine(appDir, FFmpegPath);
        }

        public bool IsFFmpegAvailable() {
            string resolved = GetResolvedFFmpegPath();
            return !string.IsNullOrEmpty(resolved) && File.Exists(resolved);
        }

        public string GetResolvedInstaTempDir() {
            if (!string.IsNullOrWhiteSpace(InstaTempDirectory)) {
                return InstaTempDirectory;
            }
            return Path.Combine(Path.GetTempPath(), "ComComTag_Insta");
        }

        public void FlushInstaTempDirectory() {
            try {
                string defaultTemp = Path.Combine(Path.GetTempPath(), "ComComTag_Insta");
                string dir = GetResolvedInstaTempDir();
                if (!Directory.Exists(dir)) return;

                bool isDefaultDir = string.IsNullOrWhiteSpace(InstaTempDirectory) ||
                    string.Equals(Path.GetFullPath(dir).TrimEnd('\\', '/'), Path.GetFullPath(defaultTemp).TrimEnd('\\', '/'), StringComparison.OrdinalIgnoreCase);

                if (isDefaultDir) {
                    // Safe to wipe entire ComComTag-exclusive temp folder
                    Directory.Delete(dir, true);
                } else {
                    // Custom user folder (e.g. C:\test): Never delete root or non-ComComTag user files!
                    // Only purge subfolders created for Instagram shortcodes containing slide_* files
                    foreach (var subDir in Directory.GetDirectories(dir)) {
                        var slideFiles = Directory.GetFiles(subDir, "slide_*.*");
                        if (slideFiles.Length > 0) {
                            try { Directory.Delete(subDir, true); } catch { }
                        }
                    }
                    // Clean any loose slide_* files in the root of the custom directory
                    foreach (var file in Directory.GetFiles(dir, "slide_*.*")) {
                        try { File.Delete(file); } catch { }
                    }
                }
            } catch { }
        }

        public void AddArtistIfNew(string artist) {
            if (string.IsNullOrWhiteSpace(artist)) return;
            string trimmed = artist.Trim();
            if (!Artists.Contains(trimmed)) {
                Artists.Add(trimmed);
                Save();
            }
        }

        public void AddLocationIfNew(string location) {
            if (string.IsNullOrWhiteSpace(location)) return;
            string trimmed = location.Trim();
            if (!Locations.Contains(trimmed)) {
                Locations.Add(trimmed);
                Save();
            }
        }

        public string FormatFilename(DateTime date, string artist, string show, string location, string trackId = "") {
            string df = !string.IsNullOrWhiteSpace(DateFormat) ? DateFormat : "yyyy-MM-dd";
            string dateStr;
            try {
                dateStr = date.ToString(df);
            } catch {
                dateStr = date.ToString("yyyy-MM-dd");
            }

            string pattern = !string.IsNullOrWhiteSpace(FilenamePattern) ? FilenamePattern : "{Date} - {Artist} - {Show} - {Location}";
            
            // If trackId is supplied and pattern doesn't contain {Track}, append ({Track})
            if (!string.IsNullOrWhiteSpace(trackId) && !pattern.IndexOf("{track}", StringComparison.OrdinalIgnoreCase).Equals(-1)) {
                // Pattern will handle track
            } else if (!string.IsNullOrWhiteSpace(trackId)) {
                pattern += " ({Track})";
            }

            // Replace placeholders with case sensitivity handling
            string result = pattern;

            // Date
            result = ReplacePlaceholder(result, "Date", dateStr);

            // Artist
            result = ReplacePlaceholder(result, "Artist", artist);

            // Show
            result = ReplacePlaceholder(result, "Show", show);

            // Location / Venue
            result = ReplacePlaceholder(result, "Location", location);
            result = ReplacePlaceholder(result, "Venue", location);

            // Track
            result = ReplacePlaceholder(result, "Track", trackId);

            // Clean up missing optional fields and trailing/duplicate separators
            // 1. Remove empty parentheses e.g. "()"
            result = Regex.Replace(result, @"\(\s*\)", "");

            // 2. Collapse multiple separators like " -  - " or " _  _ " or " - -"
            result = Regex.Replace(result, @"(\s*-\s*){2,}", " - ");
            result = Regex.Replace(result, @"(\s*_\s*){2,}", " _ ");
            result = Regex.Replace(result, @"(\s*\.\s*){2,}", " . ");

            // 3. Trim leading/trailing separators
            result = result.Trim();
            result = Regex.Replace(result, @"^(\s*-\s*|\s*_\s*|\s*\.\s*)+", "");
            result = Regex.Replace(result, @"(\s*-\s*|\s*_\s*|\s*\.\s*)+$", "");

            return result.Trim();
        }

        private string ReplacePlaceholder(string template, string placeholderName, string value) {
            if (string.IsNullOrEmpty(template)) return "";
            string val = value ?? "";

            // {PLACEHOLDER} -> ALL CAPS
            string upperPlaceholder = "{" + placeholderName.ToUpperInvariant() + "}";
            if (template.Contains(upperPlaceholder)) {
                template = template.Replace(upperPlaceholder, val.ToUpperInvariant());
            }

            // {placeholder} -> lowercase
            string lowerPlaceholder = "{" + placeholderName.ToLowerInvariant() + "}";
            if (template.Contains(lowerPlaceholder)) {
                template = template.Replace(lowerPlaceholder, val.ToLowerInvariant());
            }

            // {Placeholder} or any other case -> as provided
            var regex = new Regex(@"\{" + placeholderName + @"\}", RegexOptions.IgnoreCase);
            template = regex.Replace(template, val);

            return template;
        }

        public bool ParseFromFilename(string filename, out DateTime? date, out string artist, out string show, out string location) {
            date = null;
            artist = "";
            show = "";
            location = "";

            if (string.IsNullOrWhiteSpace(filename)) return false;
            string baseName = Path.GetFileNameWithoutExtension(filename).Trim();

            // 1. First attempt matching using the active FilenamePattern
            try {
                string pattern = !string.IsNullOrWhiteSpace(FilenamePattern) ? FilenamePattern : "{Date} - {Artist} - {Show} - {Location}";
                
                // Replace placeholders with unique GUID tokens before escaping literals
                string dToken = "___CC_DATE___";
                string aToken = "___CC_ARTIST___";
                string sToken = "___CC_SHOW___";
                string lToken = "___CC_LOCATION___";
                string tToken = "___CC_TRACK___";

                string p = pattern;
                p = Regex.Replace(p, @"\{Date\}", dToken, RegexOptions.IgnoreCase);
                p = Regex.Replace(p, @"\{Artist\}", aToken, RegexOptions.IgnoreCase);
                p = Regex.Replace(p, @"\{Show\}", sToken, RegexOptions.IgnoreCase);
                p = Regex.Replace(p, @"\{Location\}", lToken, RegexOptions.IgnoreCase);
                p = Regex.Replace(p, @"\{Venue\}", lToken, RegexOptions.IgnoreCase);
                p = Regex.Replace(p, @"\{Track\}", tToken, RegexOptions.IgnoreCase);

                // Escape literals (like -, (, ), [, ], ., etc.)
                p = Regex.Escape(p);

                // Replace tokens with non-greedy named capture groups
                p = p.Replace(dToken, @"(?<Date>.+?)");
                p = p.Replace(aToken, @"(?<Artist>.+?)");
                p = p.Replace(sToken, @"(?<Show>.+?)");
                p = p.Replace(lToken, @"(?<Location>.+?)");
                p = p.Replace(tToken, @"(?<Track>.+?)");

                string regexPattern = "^" + p + @"(?:\s*\((?<ExtraTrack>[^()]+)\))?$";

                var match = Regex.Match(baseName, regexPattern, RegexOptions.IgnoreCase);
                if (match.Success) {
                    if (match.Groups["Date"].Success) {
                        DateTime d;
                        if (DateTime.TryParse(match.Groups["Date"].Value.Trim(), out d)) {
                            date = d;
                        }
                    }
                    if (match.Groups["Artist"].Success) artist = match.Groups["Artist"].Value.Trim();
                    if (match.Groups["Show"].Success) show = match.Groups["Show"].Value.Trim();
                    if (match.Groups["Location"].Success) location = match.Groups["Location"].Value.Trim();
                    return true;
                }
            } catch { }

            // 2. Fallback heuristic: Voice Recorder Regex "Voice YYMMDD_"
            var vMatch = Regex.Match(filename, @"Voice\s+(?<YY>\d{2})(?<MM>\d{2})(?<DD>\d{2})_");
            if (vMatch.Success) {
                int year = int.Parse("20" + vMatch.Groups["YY"].Value);
                int month = int.Parse(vMatch.Groups["MM"].Value);
                int day = int.Parse(vMatch.Groups["DD"].Value);
                try {
                    date = new DateTime(year, month, day);
                    return true;
                } catch { }
            }

            // 3. Fallback heuristic: Hyphen-delimited splitting "Date - Artist - Show - Location"
            var fparts = baseName.Split(new[] { " - " }, StringSplitOptions.None);
            if (fparts.Length >= 4) {
                DateTime d;
                if (DateTime.TryParse(fparts[0], out d)) date = d;
                artist = fparts[1].Trim();
                show = fparts[2].Trim();
                string loc = fparts[3].Trim();
                var locMatch = Regex.Match(loc, @"^(?<locName>.*?)\s*\([^()]+\)$");
                if (locMatch.Success) loc = locMatch.Groups["locName"].Value.Trim();
                location = loc;
                return true;
            } else if (fparts.Length >= 3) {
                DateTime d;
                if (DateTime.TryParse(fparts[0], out d)) date = d;
                artist = fparts[1].Trim();
                string loc = fparts[2].Trim();
                var locMatch = Regex.Match(loc, @"^(?<locName>.*?)\s*\([^()]+\)$");
                if (locMatch.Success) loc = locMatch.Groups["locName"].Value.Trim();
                location = loc;
                return true;
            }

            return false;
        }

        public void Load() {
            if (!File.Exists(_settingsFile)) {
                Save(); // Create default
                return;
            }

            var lines = File.ReadAllLines(_settingsFile);
            var loadedLocations = new List<string>();
            var loadedArtists = new List<string>();

            foreach (var line in lines) {
                var trimmed = line.Trim();
                if (trimmed.StartsWith("DefaultDirectory=")) {
                    DefaultDirectory = trimmed.Substring("DefaultDirectory=".Length);
                } else if (trimmed.StartsWith("Location=")) {
                    loadedLocations.Add(trimmed.Substring("Location=".Length));
                } else if (trimmed.StartsWith("Artist=")) {
                    loadedArtists.Add(trimmed.Substring("Artist=".Length));
                } else if (trimmed.StartsWith("FFmpegPath=")) {
                    FFmpegPath = trimmed.Substring("FFmpegPath=".Length);
                } else if (trimmed.StartsWith("DefaultBitrate=")) {
                    DefaultBitrate = trimmed.Substring("DefaultBitrate=".Length);
                } else if (trimmed.StartsWith("InstaTempDir=")) {
                    InstaTempDirectory = trimmed.Substring("InstaTempDir=".Length);
                } else if (trimmed.StartsWith("FlushTempOnExit=")) {
                    bool val;
                    if (bool.TryParse(trimmed.Substring("FlushTempOnExit=".Length), out val)) FlushTempOnExit = val;
                } else if (trimmed.StartsWith("AutoAddArtists=")) {
                    bool val;
                    if (bool.TryParse(trimmed.Substring("AutoAddArtists=".Length), out val)) AutoAddArtists = val;
                } else if (trimmed.StartsWith("AutoAddLocations=")) {
                    bool val;
                    if (bool.TryParse(trimmed.Substring("AutoAddLocations=".Length), out val)) AutoAddLocations = val;
                } else if (trimmed.StartsWith("UseFirstArtistAsAlbumArtist=")) {
                    bool val;
                    if (bool.TryParse(trimmed.Substring("UseFirstArtistAsAlbumArtist=".Length), out val)) UseFirstArtistAsAlbumArtist = val;
                } else if (trimmed.StartsWith("GroupAdditionalArtists=")) {
                    bool val;
                    if (bool.TryParse(trimmed.Substring("GroupAdditionalArtists=".Length), out val)) GroupAdditionalArtists = val;
                } else if (trimmed.StartsWith("GroupAdditionalArtistsThreshold=")) {
                    int thresh;
                    if (int.TryParse(trimmed.Substring("GroupAdditionalArtistsThreshold=".Length), out thresh)) GroupAdditionalArtistsThreshold = thresh;
                } else if (trimmed.StartsWith("GroupAdditionalArtistsText=")) {
                    GroupAdditionalArtistsText = trimmed.Substring("GroupAdditionalArtistsText=".Length);
                } else if (trimmed.StartsWith("DefaultM4bArtist=")) {
                    DefaultM4bArtist = trimmed.Substring("DefaultM4bArtist=".Length);
                } else if (trimmed.StartsWith("DefaultM4bArtistVarious=")) {
                    bool val;
                    if (bool.TryParse(trimmed.Substring("DefaultM4bArtistVarious=".Length), out val)) {
                        DefaultM4bArtist = val ? "Various Artists" : "";
                    }
                } else if (trimmed.StartsWith("AppendChapterDuration=")) {
                    bool val;
                    if (bool.TryParse(trimmed.Substring("AppendChapterDuration=".Length), out val)) AppendChapterDuration = val;
                } else if (trimmed.StartsWith("ClearErasesAllTags=")) {
                    bool val;
                    if (bool.TryParse(trimmed.Substring("ClearErasesAllTags=".Length), out val)) ClearErasesAllTags = val;
                } else if (trimmed.StartsWith("Theme=")) {
                    Theme = trimmed.Substring("Theme=".Length);
                } else if (trimmed.StartsWith("DateFormat=")) {
                    DateFormat = trimmed.Substring("DateFormat=".Length);
                } else if (trimmed.StartsWith("FilenamePattern=")) {
                    FilenamePattern = trimmed.Substring("FilenamePattern=".Length);
                } else if (trimmed.StartsWith("PlaygroundDate=")) {
                    PlaygroundDate = trimmed.Substring("PlaygroundDate=".Length);
                } else if (trimmed.StartsWith("PlaygroundArtist=")) {
                    PlaygroundArtist = trimmed.Substring("PlaygroundArtist=".Length);
                } else if (trimmed.StartsWith("PlaygroundShow=")) {
                    PlaygroundShow = trimmed.Substring("PlaygroundShow=".Length);
                } else if (trimmed.StartsWith("PlaygroundLocation=")) {
                    PlaygroundLocation = trimmed.Substring("PlaygroundLocation=".Length);
                }
            }

            if (loadedLocations.Count > 0) {
                Locations = loadedLocations;
            } else if (Locations.Count == 0) {
                Locations.Add("Default Location");
            }

            if (loadedArtists.Count > 0) {
                Artists = loadedArtists;
            } else if (Artists.Count == 0) {
                Artists.Add("Various Artists");
            }

            if (string.IsNullOrWhiteSpace(Theme)) Theme = "System";
            if (string.IsNullOrWhiteSpace(DateFormat)) DateFormat = "yyyy-MM-dd";
            if (string.IsNullOrWhiteSpace(FilenamePattern)) FilenamePattern = "{Date} - {Artist} - {Show} - {Location}";
            if (GroupAdditionalArtistsThreshold < 2) GroupAdditionalArtistsThreshold = 3;
            if (string.IsNullOrWhiteSpace(GroupAdditionalArtistsText)) GroupAdditionalArtistsText = "and friends";
            if (DefaultM4bArtist == null) DefaultM4bArtist = "Various Artists";
            if (string.IsNullOrWhiteSpace(PlaygroundDate)) PlaygroundDate = "1999-06-13";
            if (string.IsNullOrWhiteSpace(PlaygroundArtist)) PlaygroundArtist = "Wally Bazoom";
            if (string.IsNullOrWhiteSpace(PlaygroundShow)) PlaygroundShow = "Smile Time";
            if (string.IsNullOrWhiteSpace(PlaygroundLocation)) PlaygroundLocation = "The Aigburth Arms, Liverpool";
        }

        public void Save() {
            using (var writer = new StreamWriter(_settingsFile)) {
                writer.WriteLine(string.Format("DefaultDirectory={0}", DefaultDirectory));
                writer.WriteLine(string.Format("FFmpegPath={0}", FFmpegPath));
                writer.WriteLine(string.Format("DefaultBitrate={0}", DefaultBitrate));
                writer.WriteLine(string.Format("InstaTempDir={0}", InstaTempDirectory));
                writer.WriteLine(string.Format("FlushTempOnExit={0}", FlushTempOnExit));
                writer.WriteLine(string.Format("AutoAddArtists={0}", AutoAddArtists));
                writer.WriteLine(string.Format("AutoAddLocations={0}", AutoAddLocations));
                writer.WriteLine(string.Format("UseFirstArtistAsAlbumArtist={0}", UseFirstArtistAsAlbumArtist));
                writer.WriteLine(string.Format("GroupAdditionalArtists={0}", GroupAdditionalArtists));
                writer.WriteLine(string.Format("GroupAdditionalArtistsThreshold={0}", GroupAdditionalArtistsThreshold));
                writer.WriteLine(string.Format("GroupAdditionalArtistsText={0}", GroupAdditionalArtistsText));
                writer.WriteLine(string.Format("DefaultM4bArtist={0}", DefaultM4bArtist));
                writer.WriteLine(string.Format("AppendChapterDuration={0}", AppendChapterDuration));
                writer.WriteLine(string.Format("ClearErasesAllTags={0}", ClearErasesAllTags));
                writer.WriteLine(string.Format("Theme={0}", Theme));
                writer.WriteLine(string.Format("DateFormat={0}", DateFormat));
                writer.WriteLine(string.Format("FilenamePattern={0}", FilenamePattern));
                writer.WriteLine(string.Format("PlaygroundDate={0}", PlaygroundDate));
                writer.WriteLine(string.Format("PlaygroundArtist={0}", PlaygroundArtist));
                writer.WriteLine(string.Format("PlaygroundShow={0}", PlaygroundShow));
                writer.WriteLine(string.Format("PlaygroundLocation={0}", PlaygroundLocation));
                foreach (var art in Artists) {
                    writer.WriteLine(string.Format("Artist={0}", art));
                }
                foreach (var loc in Locations) {
                    writer.WriteLine(string.Format("Location={0}", loc));
                }
            }
        }
    }
}
