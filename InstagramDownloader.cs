using System;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Collections.Generic;

namespace ComComTag {
    public static class InstagramDownloader {
        private const string IG_APP_ID = "936619743392459";
        private const string USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36";
        private const string DOC_ID = "27128499623469141";

        static InstagramDownloader() {
            // Enable modern TLS 1.2 / TLS 1.3 protocols on .NET Framework 4.x
            try {
                ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072 | (SecurityProtocolType)768 | SecurityProtocolType.Tls;
            } catch { }
        }

        public static string ExtractShortcode(string input) {
            if (string.IsNullOrWhiteSpace(input)) return "";
            string trimmed = input.Trim();
            var match = Regex.Match(trimmed, @"(?:instagram\.com/(?:p|reel|tv)/|/p/|/reel/)(?<codecode>[A-Za-z0-9_-]+)", RegexOptions.IgnoreCase);
            if (match.Success) {
                return match.Groups["codecode"].Value;
            }
            // If the user pasted just the shortcode itself (e.g., "Dcon0V8m71n")
            if (Regex.IsMatch(trimmed, @"^[A-Za-z0-9_-]{5,}$")) {
                return trimmed;
            }
            return trimmed.Trim('/').Split('/').Length > 0 ? trimmed.Trim('/').Split('/')[trimmed.Trim('/').Split('/').Length - 1] : trimmed;
        }

        public static List<string> DownloadPostImages(string urlOrShortcode, string targetTempDir) {
            string shortcode = ExtractShortcode(urlOrShortcode);
            if (string.IsNullOrEmpty(shortcode)) {
                throw new Exception("Invalid Instagram post URL or shortcode.");
            }

            string targetDir = Path.Combine(
                !string.IsNullOrWhiteSpace(targetTempDir) ? targetTempDir : Path.Combine(Path.GetTempPath(), "ComComTag_Insta"),
                shortcode
            );

            if (!Directory.Exists(targetDir)) {
                Directory.CreateDirectory(targetDir);
            }

            // Always fetch fresh image URLs
            List<string> imageUrls = FetchImageUrls(shortcode);
            if (imageUrls.Count == 0) {
                // If existing local images exist, fallback to them
                var existing = GetLocalImages(targetDir);
                if (existing.Count > 0) return existing;
                throw new Exception("Could not retrieve images for this Instagram post. The post may be private or removed.");
            }

            // Clear old cache in this targetDir to ensure latest uncropped images
            try {
                foreach (var oldFile in Directory.GetFiles(targetDir, "slide_*.*")) {
                    File.Delete(oldFile);
                }
            } catch { }

            List<string> downloadedFiles = new List<string>();
            using (var client = new TimeoutWebClient(10000)) {
                client.Headers.Add("User-Agent", USER_AGENT);
                client.Headers.Add("Accept", "image/avif,image/webp,image/apng,image/svg+xml,image/*,*/*;q=0.8");
                client.Headers.Add("Referer", "https://www.instagram.com/");

                for (int i = 0; i < imageUrls.Count; i++) {
                    string imgUrl = imageUrls[i];
                    string extension = ".jpg";
                    if (imgUrl.Contains(".png")) extension = ".png";
                    else if (imgUrl.Contains(".webp")) extension = ".webp";

                    string destFile = Path.Combine(targetDir, string.Format("slide_{0:D2}{1}", i + 1, extension));
                    
                    try {
                        client.DownloadFile(imgUrl, destFile);
                        if (File.Exists(destFile) && new FileInfo(destFile).Length > 0) {
                            downloadedFiles.Add(destFile);
                        }
                    } catch (Exception ex) {
                        System.Diagnostics.Debug.WriteLine("Failed to download slide " + i + ": " + ex.Message);
                    }
                }
            }

            if (downloadedFiles.Count == 0) {
                var existing = GetLocalImages(targetDir);
                if (existing.Count > 0) return existing;
                throw new Exception("Failed to save downloaded images from Instagram.");
            }

            return downloadedFiles;
        }

        private static List<string> FetchImageUrls(string shortcode) {
            var urls = new List<string>();
            var seen = new HashSet<string>();

            // Strategy 1: PolarisPostRootQuery GraphQL API (Extracts all uncropped 1440p / 1080p carousel slides)
            try {
                ExtractViaGraphQL(shortcode, urls, seen);
            } catch (Exception ex) {
                System.Diagnostics.Debug.WriteLine("GraphQL Strategy failed: " + ex.Message);
            }

            // Strategy 2: Official oEmbed API (Guaranteed uncropped single image / thumbnail)
            if (urls.Count == 0) {
                try {
                    ExtractViaOEmbed(shortcode, urls, seen);
                } catch (Exception ex) {
                    System.Diagnostics.Debug.WriteLine("oEmbed Strategy failed: " + ex.Message);
                }
            }

            // Strategy 3: Facebookbot / Crawler HTML (Fallback)
            if (urls.Count == 0) {
                try {
                    ExtractViaCrawlerHtml(shortcode, urls, seen);
                } catch (Exception ex) {
                    System.Diagnostics.Debug.WriteLine("Crawler Strategy failed: " + ex.Message);
                }
            }

            return urls;
        }

        private static void ExtractViaGraphQL(string shortcode, List<string> urls, HashSet<string> seen) {
            var cookieJar = new CookieContainer();
            string csrfToken = "";

            // 1. Visit Instagram homepage to establish session and get CSRF token
            var initReq = (HttpWebRequest)WebRequest.Create("https://www.instagram.com/");
            initReq.CookieContainer = cookieJar;
            initReq.UserAgent = USER_AGENT;
            initReq.Headers.Add("Accept-Language", "en-US,en;q=0.9");
            initReq.Timeout = 8000;

            using (var resp = (HttpWebResponse)initReq.GetResponse()) {
                foreach (Cookie c in resp.Cookies) {
                    if (c.Name == "csrftoken") csrfToken = c.Value;
                }
            }

            // 2. Query GraphQL endpoint with PolarisPostRootQuery
            var req = (HttpWebRequest)WebRequest.Create("https://www.instagram.com/graphql/query");
            req.Method = "POST";
            req.CookieContainer = cookieJar;
            req.ContentType = "application/x-www-form-urlencoded";
            req.UserAgent = USER_AGENT;
            req.Headers.Add("X-IG-App-ID", IG_APP_ID);
            req.Headers.Add("X-FB-Friendly-Name", "PolarisPostRootQuery");
            if (!string.IsNullOrEmpty(csrfToken)) req.Headers.Add("X-CSRFToken", csrfToken);
            req.Headers.Add("X-Requested-With", "XMLHttpRequest");
            req.Headers.Add("Sec-Fetch-Mode", "cors");
            req.Headers.Add("Sec-Fetch-Site", "same-origin");
            req.Headers.Add("Origin", "https://www.instagram.com");
            req.Referer = "https://www.instagram.com/p/" + shortcode + "/";
            req.Timeout = 10000;

            string vars = "{\"shortcode\":\"" + shortcode + "\",\"__relay_internal__pv__PolarisAIGMMediaWebLabelEnabledrelayprovider\":false}";
            string postData = "doc_id=" + DOC_ID + "&variables=" + Uri.EscapeDataString(vars);
            byte[] data = Encoding.UTF8.GetBytes(postData);
            req.ContentLength = data.Length;

            using (var stream = req.GetRequestStream()) {
                stream.Write(data, 0, data.Length);
            }

            using (var resp = (HttpWebResponse)req.GetResponse()) {
                using (var r = new StreamReader(resp.GetResponseStream(), Encoding.UTF8)) {
                    string json = r.ReadToEnd();
                    
                    // Match candidates inside image_versions2 (each carousel item has one)
                    var matches = Regex.Matches(json, @"""image_versions2""\s*:\s*\{\s*""candidates""\s*:\s*\[\s*\{\s*""url""\s*:\s*""([^""]+)""");
                    foreach (Match m in matches) {
                        string rawUrl = Regex.Unescape(m.Groups[1].Value.Replace("\\/", "/"));
                        rawUrl = WebUtility.HtmlDecode(rawUrl);
                        AddUrlIfValid(rawUrl, urls, seen);
                    }

                    // Fallback to display_url if candidates was empty
                    if (urls.Count == 0) {
                        var displayMatches = Regex.Matches(json, @"""display_url""\s*:\s*""([^""]+)""");
                        foreach (Match m in displayMatches) {
                            string rawUrl = Regex.Unescape(m.Groups[1].Value.Replace("\\/", "/"));
                            rawUrl = WebUtility.HtmlDecode(rawUrl);
                            AddUrlIfValid(rawUrl, urls, seen);
                        }
                    }
                }
            }
        }

        private static void ExtractViaOEmbed(string shortcode, List<string> urls, HashSet<string> seen) {
            using (var wc = new TimeoutWebClient(8000)) {
                wc.Headers.Add("User-Agent", USER_AGENT);
                string json = wc.DownloadString("https://www.instagram.com/api/v1/oembed/?url=https://www.instagram.com/p/" + shortcode + "/");
                var match = Regex.Match(json, @"""thumbnail_url""\s*:\s*""([^""]+)""");
                if (match.Success) {
                    string rawUrl = Regex.Unescape(match.Groups[1].Value.Replace("\\/", "/"));
                    rawUrl = WebUtility.HtmlDecode(rawUrl);
                    AddUrlIfValid(rawUrl, urls, seen);
                }
            }
        }

        private static void ExtractViaCrawlerHtml(string shortcode, List<string> urls, HashSet<string> seen) {
            using (var wc = new TimeoutWebClient(8000)) {
                wc.Headers.Add("User-Agent", "facebookexternalhit/1.1 (+http://www.facebook.com/externalhit_uatext.php)");
                string html = wc.DownloadString("https://www.instagram.com/p/" + shortcode + "/");

                var matches = Regex.Matches(html, @"https?:\\?/\\?/[^""'\s<>]+\.(?:heic|jpg|jpeg|png|webp)\?[^""'\s<>]*oh=[A-Za-z0-9_-]+[^""'\s<>]*", RegexOptions.IgnoreCase);
                foreach (Match m in matches) {
                    string u = Regex.Unescape(m.Value.Replace("\\/", "/"));
                    u = WebUtility.HtmlDecode(u).Trim('"', '\'', '\\', ';', ',');
                    AddUrlIfValid(u, urls, seen);
                }
            }
        }

        private static void AddUrlIfValid(string url, List<string> urls, HashSet<string> seen) {
            if (string.IsNullOrWhiteSpace(url)) return;
            url = url.Trim();
            if (!url.StartsWith("http://") && !url.StartsWith("https://")) return;
            if (url.Contains("rsrc.php") || url.Contains("static.cdninstagram.com")) return;
            
            if (!seen.Contains(url)) {
                seen.Add(url);
                urls.Add(url);
            }
        }

        private static List<string> GetLocalImages(string directory) {
            var list = new List<string>();
            if (Directory.Exists(directory)) {
                var files = Directory.GetFiles(directory, "*.*");
                foreach (var f in files) {
                    string ext = Path.GetExtension(f).ToLowerInvariant();
                    if (ext == ".jpg" || ext == ".jpeg" || ext == ".png" || ext == ".webp") {
                        list.Add(f);
                    }
                }
                list.Sort();
            }
            return list;
        }

        public class TimeoutWebClient : WebClient {
            public int TimeoutMs { get; set; }

            public TimeoutWebClient(int timeoutMs = 8000) {
                TimeoutMs = timeoutMs;
            }

            protected override WebRequest GetWebRequest(Uri address) {
                WebRequest request = base.GetWebRequest(address);
                if (request != null) {
                    request.Timeout = TimeoutMs;
                }
                return request;
            }
        }
    }
}
