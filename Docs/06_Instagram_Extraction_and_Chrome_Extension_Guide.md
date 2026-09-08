# Instagram Image Extraction & Chrome Extension Architecture Guide

A comprehensive technical blueprint explaining how to extract uncropped, original-resolution images (single posts & multi-image carousels) from Instagram without external dependencies, along with an end-to-end guide for building a **Google Chrome Extension (Manifest V3)** for right-click image downloading and clipboard copying.

---

## 1. Instagram Media & CDN Architecture

### 1.1 Shortcode Identification
Every Instagram post, reel, or carousel has a unique base64-style identifier called a **shortcode**:
- Post: `https://www.instagram.com/p/Dcon0V8m71n/` $\rightarrow$ Shortcode: `Dcon0V8m71n`
- Reel: `https://www.instagram.com/reel/C8xyz123abc/` $\rightarrow$ Shortcode: `C8xyz123abc`

Regex for extracting shortcodes across URLs:
```javascript
const shortcodeRegex = /(?:instagram\.com\/(?:p|reel|tv)\/|^\s*)([A-Za-z0-9_-]+)/i;
const match = url.match(shortcodeRegex);
const shortcode = match ? match[1] : null;
```

---

### 1.2 The CDN HMAC Signature & The "Crop Trap"
Instagram serves images via Meta's content delivery networks (`scontent*.cdninstagram.com` or `fbcdn.net`).

#### The Problem:
1. **URL Parameters**:
   - `stp=dst-jpg_e35_s1080x1080_sh0.08` (Uncropped square/original aspect ratio).
   - `stp=c0.134.1080.1080a_dst-jpg_e35_s640x640_sh0.08` (Cropped thumbnail).
2. **HMAC Signing**:
   - Instagram signs the *entire* URL string with cryptographic parameters (`&_nc_sid=...&_nc_ohc=...&oh=...&oe=...`).
   - If you naively strip or modify `stp=c...` to uncrop an image, the HMAC signature verification fails immediately with **`403 Forbidden` / `URL signature expired`**.

#### The Solution:
Rather than modifying a cropped URL, you must obtain the **signed URL of the original resolution resource** directly from Instagram's GraphQL data layer (`display_resources` array or `display_url`).

---

## 2. Multi-Tier Extraction Strategies

### Summary Architecture:
```
+-----------------------------------------------------------------------+
| Strategy 1: PolarisPostRootQuery GraphQL API (Tier 1 - Primary)       |
| - Retrieves 100% full-resolution (1440p / 1080p) uncropped slides     |
| - Supports all carousel items with original dimensions & aspect ratio |
+-----------------------------------+-----------------------------------+
                                    | (if unauthenticated / blocked)
                                    v
+-----------------------------------------------------------------------+
| Strategy 2: Official Meta oEmbed API (Tier 2 - Uncropped Single)      |
| - Queries /api/v1/oembed/?url=...                                     |
| - Returns signed uncropped primary thumbnail                          |
+-----------------------------------+-----------------------------------+
                                    | (if fallback required)
                                    v
+-----------------------------------------------------------------------+
| Strategy 3: Facebookbot Crawler Scraping (Tier 3 - Crawler Fallback)  |
| - Emulates facebookexternalhit/1.1 crawler                            |
| - Extracts signed og:image metadata tags                              |
+-----------------------------------------------------------------------+
```

---

### 2.1 Strategy 1: `PolarisPostRootQuery` GraphQL Engine (Highest Quality)

This is the exact API Instagram's web application uses to hydrate post modals.

- **Endpoint**: `POST https://www.instagram.com/graphql/query`
- **Headers**:
  ```http
  Content-Type: application/x-www-form-urlencoded
  X-IG-App-ID: 936619743392459
  X-CSRFToken: <csrftoken-from-cookies>
  X-ASBD-ID: 129477
  X-Requested-With: XMLHttpRequest
  User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36
  Referer: https://www.instagram.com/p/{shortcode}/
  ```
- **Payload (`doc_id`)**:
  - `doc_id`: `27128499623469141` *(PolarisPostRootQuery)* or `8845758582119845`
  - `variables`:
    ```json
    {
      "shortcode": "Dcon0V8m71n",
      "fetch_tagged_user_count": null,
      "hoisted_comment_id": null,
      "hoisted_feature_mention_id": null
    }
    ```

#### JSON Response Schema:
```json
{
  "data": {
    "xdt_shortcode_media": {
      "__typename": "XDTGraphSidecar",
      "id": "31234567890",
      "shortcode": "Dcon0V8m71n",
      "dimensions": { "height": 1440, "width": 1440 },
      "display_url": "https://scontent.cdninstagram.com/.../orig.jpg?...",
      "edge_sidecar_to_children": {
        "edges": [
          {
            "node": {
              "id": "31234567891",
              "display_url": "https://scontent.cdninstagram.com/.../slide_1_1440x1440.jpg?...",
              "display_resources": [
                { "src": "https://scontent.../slide_1_640.jpg", "config_width": 640 },
                { "src": "https://scontent.../slide_1_750.jpg", "config_width": 750 },
                { "src": "https://scontent.../slide_1_1080.jpg", "config_width": 1080 },
                { "src": "https://scontent.../slide_1_1440.jpg", "config_width": 1440 }
              ]
            }
          },
          {
            "node": {
              "id": "31234567892",
              "display_url": "https://scontent.cdninstagram.com/.../slide_2_1440x1440.jpg?..."
            }
          }
        ]
      }
    }
  }
}
```

#### Slide Selection Algorithm:
1. Check if `xdt_shortcode_media.edge_sidecar_to_children` is present and contains items:
   - If **yes**: Iterate through `edges[]`. For each child node, select `display_resources[last].src` or fallback to `display_url`.
   - If **no**: Single image post. Select `xdt_shortcode_media.display_resources[last].src` or `display_url`.
2. This guarantees you receive the maximum resolution available (up to **1440×1440** or **1080×1350** uncropped).

---

### 2.2 Strategy 2: Official Meta oEmbed API

Publicly accessible for public posts without needing custom cookies:
- **Endpoint**: `GET https://www.instagram.com/api/v1/oembed/?url=https://www.instagram.com/p/{shortcode}/`
- **Response**:
  ```json
  {
    "title": "Post caption...",
    "author_name": "instagram_user",
    "thumbnail_url": "https://scontent.cdninstagram.com/.../uncropped.jpg?...",
    "thumbnail_width": 640,
    "thumbnail_height": 640
  }
  ```

---

### 2.3 Strategy 3: Facebookbot Crawler Scraping

Emulates the OpenGraph social sharing preview crawler:
- **Endpoint**: `GET https://www.instagram.com/p/{shortcode}/`
- **Header**: `User-Agent: facebookexternalhit/1.1 (+http://www.facebook.com/externalhit_uatext.php)`
- **HTML Extraction**:
  ```html
  <meta property="og:image" content="https://scontent.../uncropped_signed.jpg?..." />
  ```

---

## 3. Chrome Extension Blueprint (Manifest V3)

Building a Chrome extension gives you a major advantage over external scripts: **the extension runs in the user's browser with access to their active Instagram session cookies**, allowing it to extract images from private accounts they follow, age-gated posts, and carousels without any login blocks.

---

### 3.1 Extension File Structure
```
instagram-image-grabber/
├── manifest.json
├── background.js       (Service Worker - handles context menus, API calls, downloads)
├── content.js          (DOM inspection - detects active post / right-clicked image)
└── icons/
    ├── icon16.png
    ├── icon48.png
    └── icon128.png
```

---

### 3.2 `manifest.json` (Manifest V3)

```json
{
  "manifest_version": 3,
  "name": "Instagram High-Res Image Saver",
  "version": "1.0.0",
  "description": "Right-click to download or copy full-resolution uncropped Instagram images and carousels.",
  "permissions": [
    "contextMenus",
    "activeTab",
    "downloads",
    "clipboardWrite",
    "cookies"
  ],
  "host_permissions": [
    "*://*.instagram.com/*"
  ],
  "background": {
    "service_worker": "background.js"
  },
  "content_scripts": [
    {
      "matches": ["*://*.instagram.com/*"],
      "js": ["content.js"],
      "run_at": "document_idle"
    }
  ],
  "icons": {
    "16": "icons/icon16.png",
    "48": "icons/icon48.png",
    "128": "icons/icon128.png"
  }
}
```

---

### 3.3 `content.js` (DOM Interaction)

Detects the clicked element, resolves the post shortcode from the DOM or URL, and tracks active carousel slide indices:

```javascript
// content.js - DOM interaction for Instagram
let lastRightClickedElement = null;

document.addEventListener("contextmenu", (event) => {
  lastRightClickedElement = event.target;
}, true);

// Listen for messages from background service worker
chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
  if (request.action === "GET_POST_INFO") {
    const postInfo = getActivePostInfo(lastRightClickedElement);
    sendResponse(postInfo);
  }
  return true;
});

function getActivePostInfo(target) {
  // 1. Try to find shortcode from current URL (e.g. /p/SHORTCODE/)
  const urlMatch = window.location.pathname.match(/\/(?:p|reel|tv)\/([A-Za-z0-9_-]+)/);
  let shortcode = urlMatch ? urlMatch[1] : null;

  // 2. If viewing a feed/grid, find closest enclosing <article> or <a> tag
  if (!shortcode && target) {
    const article = target.closest("article");
    if (article) {
      const link = article.querySelector('a[href*="/p/"], a[href*="/reel/"]');
      if (link) {
        const linkMatch = link.href.match(/\/(?:p|reel|tv)\/([A-Za-z0-9_-]+)/);
        if (linkMatch) shortcode = linkMatch[1];
      }
    }
  }

  // 3. Detect carousel slide index if inside an active slider
  let slideIndex = 0;
  if (target) {
    const sliderContainer = target.closest("ul");
    if (sliderContainer) {
      const activeSlide = target.closest("li");
      if (activeSlide && sliderContainer.children) {
        slideIndex = Array.from(sliderContainer.children).indexOf(activeSlide);
        if (slideIndex < 0) slideIndex = 0;
      }
    }
  }

  return {
    shortcode: shortcode,
    slideIndex: slideIndex,
    pageUrl: window.location.href
  };
}
```

---

### 3.4 `background.js` (GraphQL Extractor, Downloader & Clipboard Engine)

```javascript
// background.js - Service Worker

// 1. Create Context Menu on installation
chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: "ig_download_current",
    title: "Download Full-Res Image",
    contexts: ["image", "video", "link", "page"]
  });

  chrome.contextMenus.create({
    id: "ig_copy_current",
    title: "Copy Full-Res Image to Clipboard",
    contexts: ["image", "video", "link", "page"]
  });

  chrome.contextMenus.create({
    id: "ig_download_all",
    title: "Download All Carousel Images",
    contexts: ["image", "video", "link", "page"]
  });
});

// 2. Handle Context Menu Clicks
chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (!tab || !tab.id) return;

  try {
    // Query content script for target post info
    const response = await chrome.tabs.sendMessage(tab.id, { action: "GET_POST_INFO" });
    if (!response || !response.shortcode) {
      console.warn("Could not determine Instagram post shortcode.");
      return;
    }

    const { shortcode, slideIndex } = response;
    const images = await fetchInstagramImages(shortcode);

    if (!images || images.length === 0) {
      console.error("No high-resolution images found for post:", shortcode);
      return;
    }

    const selectedUrl = images[slideIndex] || images[0];

    if (info.menuItemId === "ig_download_current") {
      downloadImage(selectedUrl, `${shortcode}_slide_${slideIndex + 1}.jpg`);
    } else if (info.menuItemId === "ig_copy_current") {
      await copyImageToClipboard(tab.id, selectedUrl);
    } else if (info.menuItemId === "ig_download_all") {
      images.forEach((imgUrl, idx) => {
        downloadImage(imgUrl, `${shortcode}_slide_${idx + 1}.jpg`);
      });
    }
  } catch (err) {
    console.error("Error processing Instagram image action:", err);
  }
});

// 3. Fetch Full-Resolution Images via GraphQL API
async function fetchInstagramImages(shortcode) {
  // Retrieve CSRF token from cookies
  const cookie = await chrome.cookies.get({ url: "https://www.instagram.com", name: "csrftoken" });
  const csrfToken = cookie ? cookie.value : "";

  const payload = new URLSearchParams({
    doc_id: "27128499623469141",
    variables: JSON.stringify({
      shortcode: shortcode,
      fetch_tagged_user_count: null,
      hoisted_comment_id: null,
      hoisted_feature_mention_id: null
    })
  });

  const response = await fetch("https://www.instagram.com/graphql/query", {
    method: "POST",
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
      "X-IG-App-ID": "936619743392459",
      "X-CSRFToken": csrfToken,
      "X-Requested-With": "XMLHttpRequest"
    },
    body: payload.toString()
  });

  if (!response.ok) {
    throw new Error(`GraphQL query failed with status: ${response.status}`);
  }

  const json = await response.json();
  const media = json?.data?.xdt_shortcode_media;
  if (!media) return [];

  const results = [];

  // Check for Carousel (Sidecar)
  if (media.edge_sidecar_to_children?.edges?.length > 0) {
    for (const edge of media.edge_sidecar_to_children.edges) {
      const node = edge.node;
      if (node.display_resources?.length > 0) {
        // Highest resolution is the last item in display_resources
        results.push(node.display_resources[node.display_resources.length - 1].src);
      } else if (node.display_url) {
        results.push(node.display_url);
      }
    }
  } else {
    // Single image
    if (media.display_resources?.length > 0) {
      results.push(media.display_resources[media.display_resources.length - 1].src);
    } else if (media.display_url) {
      results.push(media.display_url);
    }
  }

  return results;
}

// 4. Download Helper
function downloadImage(url, filename) {
  chrome.downloads.download({
    url: url,
    filename: `Instagram/${filename}`,
    saveAs: false
  });
}

// 5. Copy to Clipboard Helper (fetches blob and writes PNG/JPEG)
async function copyImageToClipboard(tabId, imageUrl) {
  chrome.scripting.executeScript({
    target: { tabId: tabId },
    func: async (url) => {
      try {
        const res = await fetch(url);
        const blob = await res.blob();
        
        // Convert to PNG for maximum clipboard compatibility
        const img = new Image();
        img.crossOrigin = "anonymous";
        const loaded = new Promise((resolve) => { img.onload = resolve; });
        img.src = URL.createObjectURL(blob);
        await loaded;

        const canvas = document.createElement("canvas");
        canvas.width = img.naturalWidth;
        canvas.height = img.naturalHeight;
        const ctx = canvas.getContext("2d");
        ctx.drawImage(img, 0, 0);

        canvas.toBlob(async (pngBlob) => {
          await navigator.clipboard.write([
            new ClipboardItem({ "image/png": pngBlob })
          ]);
          console.log("Full-res Instagram image copied to clipboard!");
        }, "image/png");
      } catch (e) {
        console.error("Clipboard copy failed:", e);
      }
    },
    args: [imageUrl]
  });
}
```

---

## 4. Key Takeaways & Best Practices

1. **Never mutate crop parameters (`stp=c...`) directly on CDN URLs**: Always obtain the signed URL from the GraphQL `display_resources` array or `display_url`.
2. **Use the `936619743392459` App ID**: This standard Web App ID allows access to Instagram's post query endpoint.
3. **Session Re-use**: In a Chrome extension, utilizing `chrome.cookies` or existing session headers eliminates rate limits and allows retrieving images from private accounts you follow.
4. **Clipboard Compatibility**: Always draw image blobs onto a temporary `<canvas>` and write as `"image/png"` to ensure seamless pasting across all Windows and Mac software (Photoshop, Word, Discord, Paint, etc.).
