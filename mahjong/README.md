# Mahjong Game Night Refresher

Deploy these two files together at:

`https://anandsinhausa.com/mahjong/`

## Files

- `index.html` — self-contained Mahjong refresher. The in-page reference photos are embedded in the HTML.
- `mahjong-game-night-og.jpg` — lightweight 1200×630 social preview image for WhatsApp, iMessage, Facebook, LinkedIn, etc.

## Important

Keep the OG image as a **separate small JPG**. Social crawlers do not use images embedded as base64 inside the HTML for `og:image`. The HTML points to:

`https://anandsinhausa.com/mahjong/mahjong-game-night-og.jpg`

Therefore the folder on the web server/GitHub Pages must contain both files at the same level.

The HTML already contains canonical, Open Graph, and Twitter/X preview metadata for `https://anandsinhausa.com/mahjong/`.

After deployment, test the public URL in WhatsApp/iMessage. Social apps may cache an earlier preview, so a previously shared URL may take time to refresh.
