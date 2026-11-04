# Journal Club

Schedule and organizing team for the journal club run by Dr. Babak Taati's group (University of Toronto / KITE).

Live page: https://taatiteam.github.io/journal-club/

The same information is also shown on https://kimiashaban.github.io/journal-club/, which copies this repo's data every time it is rebuilt (on every push to that site, and once a day).

## How to add or change a session

Everything on the page comes from one file: [`_data/journal_club.yml`](_data/journal_club.yml). You do not need to clone anything.

1. Open `_data/journal_club.yml` on GitHub and click the pencil icon (Edit).
2. Add an entry under `upcoming:` (soonest first). Copy the commented example in the file. Only `date` and `speaker` are required; leave any other field as `""` and it will not be shown.
3. Commit directly to `main`. GitHub Pages rebuilds the site within a minute or two.
4. After the session, move the entry from `upcoming:` to `past:` (most recent first).

Dates are written as `YYYY-MM-DD` without quotes.

## Photos

Put organizer photos in `assets/img/journal_club/` and set `image: journal_club/<file>` for that person in the YAML. Any size works; the page crops to a circle. If the file is missing, the page shows initials instead.

## Meeting details

The `meeting:` block at the top of the YAML holds the time, room, format and contact. Blank fields are hidden.

## Previewing locally (optional)

```bash
bundle install
bundle exec jekyll serve
# open http://localhost:4000/journal-club/
```
