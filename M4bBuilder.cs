using System;
using System.IO;
using System.Text;
using System.Diagnostics;
using System.Collections.Generic;

namespace ComComTag {
    public static class M4bBuilder {
        public static string Build(List<string> mp3Paths, List<string> chapterTitles, string outputPath, string coverImagePath, string album, string artist, string bitrate, string ffmpegPath, Action<int> progressCallback, Action<string> logCallback) {
            
            if (!File.Exists(ffmpegPath)) {
                return "FFmpeg not found. Please ensure ffmpeg.exe is in the application folder or update settings.ini.";
            }

            if (mp3Paths.Count == 0 || mp3Paths.Count != chapterTitles.Count) return "No MP3s selected or title mismatch.";

            string tempDir = Path.Combine(Path.GetTempPath(), "ComComTagM4B");
            Directory.CreateDirectory(tempDir);

            // 1. Create concat file
            string concatFile = Path.Combine(tempDir, "files.txt");
            using (var writer = new StreamWriter(concatFile, false, new UTF8Encoding(false))) {
                foreach (string file in mp3Paths) {
                    // Escape single quotes and use forward slashes for ffmpeg concat demuxer
                    string safePath = file.Replace("\\", "/").Replace("'", "'\\''");
                    writer.WriteLine(string.Format("file '{0}'", safePath));
                }
            }

            // 2. Generate Chapters metadata
            string metadataFile = Path.Combine(tempDir, "metadata.txt");
            long currentOffsetMs = 0;
            using (var writer = new StreamWriter(metadataFile, false, new UTF8Encoding(false))) {
                writer.WriteLine(";FFMETADATA1");
                writer.WriteLine(string.Format("title={0}", album));
                writer.WriteLine(string.Format("artist={0}", artist));
                writer.WriteLine("");

                for (int i = 0; i < mp3Paths.Count; i++) {
                    string path = mp3Paths[i];
                    string title = chapterTitles[i];
                    long durationMs = GetDurationMs(path, ffmpegPath);

                    writer.WriteLine("[CHAPTER]");
                    writer.WriteLine("TIMEBASE=1/1000");
                    writer.WriteLine(string.Format("START={0}", currentOffsetMs));
                    writer.WriteLine(string.Format("END={0}", currentOffsetMs + durationMs));
                    writer.WriteLine(string.Format("title={0}", title));
                    writer.WriteLine("");

                    currentOffsetMs += durationMs;
                }
            }

            long totalDurationMs = currentOffsetMs;

            // 3. Run FFmpeg
            logCallback("Starting build...");
            
            string args = string.Format("-y -f concat -safe 0 -i \"{0}\" -i \"{1}\" ", concatFile, metadataFile);
            
            bool hasImage = !string.IsNullOrWhiteSpace(coverImagePath) && File.Exists(coverImagePath);
            if (hasImage) {
                args += string.Format("-i \"{0}\" -map 0:a -map 2:v -c:a aac -b:a {1} -c:v mjpeg -disposition:v attached_pic ", coverImagePath, bitrate);
            } else {
                args += string.Format("-map 0:a -c:a aac -b:a {0} ", bitrate);
            }

            args += string.Format("-map_metadata 1 \"{0}\"", outputPath);

            var startInfo = new ProcessStartInfo {
                FileName = ffmpegPath,
                Arguments = args,
                UseShellExecute = false,
                RedirectStandardError = true,
                RedirectStandardOutput = true,
                CreateNoWindow = true
            };

            var process = new Process { StartInfo = startInfo };
            
            StringBuilder errBuilder = new StringBuilder();
            
            process.ErrorDataReceived += (sender, e) => {
                if (e.Data != null) {
                    errBuilder.AppendLine(e.Data);
                    
                    // Look for: time=00:01:23.45
                    var match = System.Text.RegularExpressions.Regex.Match(e.Data, @"time=(\d{2}):(\d{2}):(\d{2})\.(\d{2})");
                    if (match.Success && totalDurationMs > 0) {
                        int h = int.Parse(match.Groups[1].Value);
                        int m_ = int.Parse(match.Groups[2].Value);
                        int s = int.Parse(match.Groups[3].Value);
                        int ms = int.Parse(match.Groups[4].Value) * 10;
                        long currentMs = (long)new TimeSpan(0, h, m_, s, ms).TotalMilliseconds;
                        
                        int percentage = (int)Math.Min(100, (currentMs * 100) / totalDurationMs);
                        progressCallback(percentage);
                    }
                }
            };

            logCallback("Transcoding... this may take a few minutes.");
            process.Start();
            process.BeginErrorReadLine();
            process.BeginOutputReadLine(); // Standard output isn't utilized for progress but must be drained.
            process.WaitForExit();

            string errOut = errBuilder.ToString();

            // A 161 byte file usually indicates ffmpeg created the header but failed to write any streams.
            bool isTooSmall = File.Exists(outputPath) && new FileInfo(outputPath).Length < 256;
            if (process.ExitCode != 0 || (errOut.Contains("Error while opening")) || isTooSmall) {
                return string.Format("FFmpeg Error ({0}):\n{1}", process.ExitCode, errOut);
            }

            // Cleanup
            try {
                Directory.Delete(tempDir, true);
            } catch { }

            return "Success";
        }

        private static long GetDurationMs(string path, string ffmpegPath) {
            if (string.IsNullOrEmpty(ffmpegPath) || !File.Exists(ffmpegPath)) return 0;

            var startInfo = new ProcessStartInfo {
                FileName = ffmpegPath,
                Arguments = string.Format("-i \"{0}\" -f null -", path),
                UseShellExecute = false,
                RedirectStandardError = true,
                CreateNoWindow = true
            };

            try {
                var process = Process.Start(startInfo);
                if (process == null) return 0;
                
                string output = process.StandardError.ReadToEnd();
                process.WaitForExit();

                // Look for: Duration: 00:01:23.45
                System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(output, @"Duration: (\d{2}):(\d{2}):(\d{2})\.(\d{2})");
                if (m.Success) {
                    int h = int.Parse(m.Groups[1].Value);
                    int m_ = int.Parse(m.Groups[2].Value);
                    int s = int.Parse(m.Groups[3].Value);
                    int ms = int.Parse(m.Groups[4].Value) * 10;
                    return (long)new TimeSpan(0, h, m_, s, ms).TotalMilliseconds;
                }
            } catch {
                return 0;
            }
            return 0; // fallback if failed
        }
    }
}
