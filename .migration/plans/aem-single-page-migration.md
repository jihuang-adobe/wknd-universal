# Cards-News Block Fix Plan

## Overview

Fix broken images and refine styling in the cards-news block to match the BD source site.

**Block:** cards-news
**Location:** `content/en-us.plain.html` and `blocks/cards-news/cards-news.css`

## Issue Analysis

### Broken Images
The images are broken because they reference external URLs (`https://news.bd.com/image/...` and `https://www.bd.com/content/dam/...`). The AEM dev server proxies these through its image optimization pipeline (`/image/BD_Elyra_Plus_400px.jpg?width=750&format=webply`), but the external domains don't allow cross-origin access from localhost — resulting in broken images.

**Fix:** Update the image `src` attributes in the `.plain.html` to use the full absolute URLs so they load directly, or adjust the proxy configuration.

### Styling Gaps
Compared to the source site, the cards need:
- Proper image aspect ratio and sizing
- Correct spacing and typography matching the BD News/Trending section

## Checklist

- [ ] Fix broken image URLs in `content/en-us.plain.html` (use absolute external URLs)
- [ ] Verify images load after fix
- [ ] Refine cards-news CSS if needed to match source layout
- [ ] Take screenshot to confirm visual match

---

_This plan requires Execute mode to make the file changes. Please switch to Execute mode to proceed._
