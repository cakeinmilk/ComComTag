# Instagram Carousel Scraping & Image Resolution Guide

This document details how ComComTag retrieves high-resolution, uncropped Instagram images and carousels, the technical investigation behind Meta's CDN cropping mechanisms, and how the multi-tiered resolution pipeline functions.

---

## 1. Overview of the Problem

When downloading images from Instagram posts (such as carousel posts like `https://www.instagram.com/p/Dcon0V8m71n/`), two distinct issues previously occurred:

1. **Only 1 slide was retrieved**: Traditional web scraping or OpenGraph metadata scraping only found the first slide, bypassing slides 2 through 6.
2. **The image appeared cropped (3/4 size)**: The preview and saved image had portions (such as the bottom-right quadrant) clipped off despite the original post image being square.

---

## 2. Technical Root Cause Analysis

### A. Meta CDN Crop Parameter (`stp=c...`)
When Meta generates crawler preview images (e.g. for `og:image` or social media link previews), it appends a server-side crop directive into the URL's `stp` query parameter:
```text
...stp=c288.0.864.864a_dst-jpg_e35_s640x640_tt6...
```
- `c288.0.864.864a` commands Meta's image server to crop from $X = 288\text{px}$, $Y = 0\text{px}$ with a bounding box of $864 \times 864\text{px}$.
- Because Meta's CDN URLs are signed with HMAC hashes (`oh=...` and `oe=...`), modifying this parameter manually invalidates the signature and yields a `403 Forbidden` error.

### B. Client-Side Rendering (React / Comet SSR)
Instagram hides carousel children (`carousel_media`) from raw HTML responses. Standard HTTP GET requests without session-level cookies only receive an empty JavaScript shell.

---

## 3. The Multi-Tiered Solution

ComComTag implements a three-tier pipeline in [`InstagramDownloader.cs`](../InstagramDownloader.cs):

```
+-------------------------------------------------------------------+
| Strategy 1: PolarisPostRootQuery GraphQL API (Tier 1 - Primary)   |
| - Negotiates session & CSRF token with Instagram homepage         |
| - POSTs to /graphql/query with doc_id 27128499623469141           |
| - Extracts full 1440x1440 / 1080p uncropped slides for all items  |
+---------------------------------+---------------------------------+
                                  | (if empty or restricted)
                                  v
+-------------------------------------------------------------------+
| Strategy 2: Official Meta oEmbed API (Tier 2 - Uncropped Single)  |
| - Queries /api/v1/oembed/?url=https://www.instagram.com/p/...     |
| - Retrieves signed uncropped thumbnail_url (no stp=c crop)        |
+---------------------------------+---------------------------------+
                                  | (if fallback needed)
                                  v
+-------------------------------------------------------------------+
| Strategy 3: Facebookbot Crawler Scraping (Tier 3 - Crawler Fallback) |
| - Requests page with facebookexternalhit/1.1                      |
| - Extracts signed scontent/fbcdn CDN assets                       |
+-------------------------------------------------------------------+
```

---

## 4. Test Verification & Results

All test harness scripts are kept in the isolated [`tests/`](../tests/) directory:

| Test Script | Description | Result for `Dcon0V8m71n` |
| :--- | :--- | :--- |
| [`tests/TestAllCarouselSlides.cs`](../tests/TestAllCarouselSlides.cs) | GraphQL query execution & slide downloader | **6 slides downloaded** at full **1440x1440** resolution |
| [`tests/TestOembedDownload.cs`](../tests/TestOembedDownload.cs) | Official oEmbed uncropped asset validation | **640x640 uncropped** image downloaded |
| [`tests/DownloadCompare.cs`](../tests/DownloadCompare.cs) | CDN signature & crop parameter analysis | Confirmed crop parameter origin |

### Sample Output:
```text
Carousel slides matched: 6
Slide 1: 1440x1440 (164344 bytes)
Slide 2: 1440x1440 (164344 bytes)
Slide 3: 1440x1440 (131021 bytes)
Slide 4: 1440x1440 (162296 bytes)
Slide 5: 1440x1440 (119526 bytes)
Slide 6: 1440x1440 (126272 bytes)
```

---

## 5. UI Integration

1. **Automatic Carousel Modal**:
   - When a multi-image carousel post is downloaded, [`ImageCarouselDialog.cs`](../ImageCarouselDialog.cs) automatically opens.
   - You can cycle through all downloaded 1440x1440 uncropped slides using the `< Previous` / `Next >` buttons or `Left` / `Right` keyboard arrow keys.
2. **Double-Click Carousel**:
   - Double-clicking the cover art preview box at any time will re-open the carousel selection dialog.
3. **Drag & Drop / File Browser**:
   - You can also drag and drop any local image file or click the cover box to choose an image directly from your disk.
