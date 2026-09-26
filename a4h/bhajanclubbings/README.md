# A4H Bhajan Clubbing — No-Diwali Landing Page

## Production URL

https://anandsinhausa.com/a4h/bhajanclubbings/

This is the separate no-Diwali version of the A4H Bhajan Clubbing landing page.

IMPORTANT: Do not replace or modify the existing `/a4h/bhajanclubbing/` directory or its OG image.

## Files to upload

Place these files together in:

`/a4h/bhajanclubbings/`

Required files:

- `index.html`
- `bhajan-clubbings-hero.png`
- `bhajan-clubbings-qr.png`
- `README.md`

## Create the directory

From the root of the local GitHub repository:

```bash
mkdir -p a4h/bhajanclubbings
```

Then copy the four files into:

```text
a4h/bhajanclubbings/
```

## Git commands

After the files are in the new directory:

```bash
git add a4h/bhajanclubbings/
git commit -m "Add A4H Bhajan Clubbing no-Diwali landing page"
git push
```

## Important deployment notes

- This is a NEW directory: `/a4h/bhajanclubbings/`
- Leave `/a4h/bhajanclubbing/` untouched.
- Leave the current OG image untouched.
- The QR points to:
  `https://anandsinhausa.com/a4h/bhajanclubbings/`
- The sponsorship/contact form uses the existing Google Apps Script / Google Sheet workflow.
- Test the new page after GitHub Pages deploys it.

## Post-deployment test

Open:

https://anandsinhausa.com/a4h/bhajanclubbings/

Confirm:

1. Page loads correctly on desktop and mobile.
2. Hero image loads.
3. QR code points to this same `/bhajanclubbings/` URL.
4. Ticket buttons work.
5. Videos load.
6. Sponsorship/contact modal opens.
7. Submit one test inquiry.
8. Confirm the inquiry appears in the existing Google Sheet.
9. Confirm the existing `/a4h/bhajanclubbing/` page is unchanged.
