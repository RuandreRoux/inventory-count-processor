You are helping the user add a new car listing website scraper to DreamCar ZA. Your job is to gather everything needed upfront so you can write the scraper without guessing.

Ask the following questions using the AskUserQuestion tool. Group them into at most 2 rounds so the user isn't overwhelmed.

---

**Round 1 — Core info**

Ask all of the following in a single AskUserQuestion call (use "Other" / free text for open-ended ones):

1. What is the website name and its base search URL?  
   (e.g. `https://www.example.co.za/used-cars/`)

2. Paste a real search URL for "Toyota Fortuner" on that site — page 1.  
   (Navigate to the site, search for Toyota Fortuner, copy the URL from the browser.)

3. Paste the same search but for **page 2**.  
   (Click to page 2, copy the URL — this reveals how pagination works.)

4. Does the site load listings via **server-side rendering** (content is in the page HTML / view-source) or via an **XHR/API call** (content loads after the page, visible in DevTools → Network)?  
   Options: SSR only / API/XHR / Both / Not sure

5. Paste a real **individual listing URL** (click any car, copy the URL).  
   (This shows us the URL slug format and ID structure.)

---

**Round 2 — Technical details**

After the user answers round 1, ask:

6. If it uses an API/XHR: open DevTools → Network → XHR/Fetch tab, reload the search page, and paste the **request URL** of the call that returns the listing data. Also note the HTTP status code.

7. Does that API request include any special **request headers** (Authorization, x-api-key, x-token, etc.)? If yes, paste them.

8. What **fields** does a listing card show? Tick all visible:  
   Price / Mileage / Year / Make / Model / Variant / Location / Image / Service history / Transmission / Fuel type

9. Is there any sign of **anti-bot protection** on the site? (Cloudflare challenge page, CAPTCHA, 403 errors when using DevTools to re-send requests, etc.)  
   Options: None / Cloudflare / CAPTCHA / 403 on API / Not sure

10. Any other notes — unusual URL patterns, login required, regional restrictions, anything else that seems relevant?

---

**After both rounds**, summarise what you learned in a short bullet list, then implement the scraper in `lib/scrapers/<sitename>.ts` following the exact same pattern as `lib/scrapers/carscoza.ts`:

- Export `async function scrape<SiteName>(browser, query, filters): Promise<Listing[]>`
- Use `normalizeRaw` from `./normalize`
- Log with `[SiteName]` prefix
- Support pagination up to 300 items
- Fall back to DOM parsing if the API is blocked

Then wire it into `lib/scrapers/index.ts` alongside the existing scrapers.

Do not start writing any code until both rounds of questions are complete and you have confirmed the answers with the user.
