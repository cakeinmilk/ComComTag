import sys
import instaloader
import re
import os
import tempfile

def download_post(url):
    try:
        L = instaloader.Instaloader(download_video_thumbnails=False, download_geotags=False, download_comments=False, save_metadata=False, quiet=True)
        
        # Extract shortcode
        match = re.search(r'(?:p|reel)/([^/?#&]+)', url)
        shortcode = match.group(1) if match else url.strip('/').split('/')[-1]
        
        post = instaloader.Post.from_shortcode(L.context, shortcode)
        
        # Save to temp directory
        target_dir = os.path.join(tempfile.gettempdir(), "ComComTag_Insta", shortcode)
        os.makedirs(target_dir, exist_ok=True)
        
        # We need to explicitly pass dirname=target_dir since target= sets the prefix rather than the directory directly if we don't change cwd.
        # Actually instaloader targets a directory name directly.
        L.download_post(post, target=target_dir)
        
        # Find all downloaded jpgs anywhere in the target_dir tree
        if os.path.exists(target_dir):
            all_jpgs = []
            for root, dirs, files in os.walk(target_dir):
                for f in files:
                    if f.endswith('.jpg'):
                        all_jpgs.append(os.path.join(root, f))
            if all_jpgs:
                for jpg in all_jpgs:
                    print(jpg)
                return
        print("ERROR: Could not find downloaded image.", file=sys.stderr)
        sys.exit(1)
    except Exception as e:
        print(f"ERROR: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    if len(sys.argv) > 1:
        download_post(sys.argv[1])
    else:
        print("ERROR: No URL provided.")
