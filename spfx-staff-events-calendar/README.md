# Shared Staff Events Calendar (SharePoint Framework)

A shared internal staff calendar concept built on SharePoint Online for a work-experience placement project at Charles Darwin University (CDU). Staff submit events through a form, reviewers approve them, and approved events appear on a branded SharePoint page in several layouts.

The page is built around an extended version of the PnP **Events Feed** web part (`react-events` sample). I added a **hover details card**: rest the pointer on an event and a card opens with the full event details.

## What I built

### 1. Hover details card (new web part feature)

- Rest the pointer on an event card, or tab to it with the keyboard, and a details card opens beside it. **Esc** or scrolling closes it.
- The card shows the event image, category, title, date and time, location and campus, delivery mode, audience, organiser and organising unit, cost, accessibility notes, registration status and the full description.
- It also shows **Register**, **Join online**, **More information** and **Open event** links when the event has them.
- It works in every layout: agenda, grid, list, filmstrip, carousel, feature, masonry, marquee and compact list.
- For SharePoint list calendars, the extra submission fields are loaded the first time you hover over an event, then cached.
- The web part setting **Show event details on hover** (on by default) turns it on or off.

Code: [`src/components/EventHoverCard/EventHoverCard.tsx`](src/components/EventHoverCard/EventHoverCard.tsx), wired up in [`EventFeedWebPartRoot.tsx`](src/components/EventFeedWebPartRoot.tsx) and [`EventsFeedWebPart.tsx`](src/webparts/eventsFeed/EventsFeedWebPart.tsx).

### 2. Event submission and approval (SharePoint list)

An Events list with the submission fields defined in the project brief:

| Field | Type | Required |
|---|---|---|
| Event title (80 characters max) | Text | Yes |
| Start / end date and time | Date and time | Yes |
| Short description | Multiple lines | Yes |
| Category | Choice (7 categories) | Yes |
| Delivery mode | In person / Online / Hybrid | Yes |
| Campus | Choice | Yes |
| Audience | Multi-choice | Yes |
| Organising unit | Choice | Yes |
| Organiser or contact | Person | Yes |
| Registration required, link and closing date | Yes/No, link, date | No |
| Online meeting link, cost, accessibility information, more information, repeats | Various | No |
| Image | Link | No |

The categories are Academic and key dates, Student events, Staff events and meetings, Professional development, Social and wellbeing, Operational, and Community and external.

**Content approval** is turned on. A new submission stays hidden from staff until a reviewer approves it, and reviewers can reject it with a comment. The list has these views:

- Calendar (month view)
- Upcoming (agenda)
- By category
- My submissions
- Approve/reject items

See [`docs/SETUP.md`](docs/SETUP.md) for how the list and page are set up.

### 3. Staff Events page

A SharePoint communication-style page with:

- a site theme in CDU colours
- a banner
- an agenda feed with event photos and a date-range filter
- a world clock for the CDU campuses
- a countdown to the next major event
- a **Submit an event** button and links to the calendar views
- a "More ways to see what's on" section with carousel, drag-to-reorder grid and compact list layouts

## Running it locally

Requirements: Node.js 22.14 or later (but below 23), and a SharePoint Online tenant.

```bash
npm install
npm run start
```

Then open the hosted workbench:

```
https://<your-tenant>.sharepoint.com/_layouts/15/workbench.aspx?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&debug=true&noredir=true
```

You can also add `?debugManifestsFile=https%3A%2F%2Flocalhost%3A4321%2Ftemp%2Fbuild%2Fmanifests.js&loadSPFX=true` to any modern page.

To build a package for the tenant app catalog:

```bash
npm run build   # produces sharepoint/solution/*.sppkg
```

The web part needs these Microsoft Graph permissions, approved in the SharePoint admin centre: `Calendars.Read`, `GroupMember.Read.All`, `Sites.Read.All` and `MailboxSettings.Read`.

## Findings from testing

- **SharePoint list events show the correct local time.** In testing, events read from a Microsoft 365 group calendar appeared several hours early. The likely cause is the `MailboxSettings.Read` permission (used to read the user's time zone) not being approved.
- **Card photos come only from SharePoint list events** (the list's banner image field). Photos on Outlook group events are not shown.
- **Events lists open in SharePoint's classic interface.** So the month view is the classic calendar, and the image is set by editing the event rather than in the new-event form.

## Credits

This project extends the [PnP `react-events` sample](https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-events) by [João Mendes](https://github.com/joaojmendes), licensed under MIT. The original sample README is kept in [`docs/PNP-SAMPLE-README.md`](docs/PNP-SAMPLE-README.md).

My additions are the hover details card, the web part setting, and the SharePoint list, approval and page design for the staff calendar concept.

**Author:** MD Rouful Alam Majumder, Master of Cyber Security student at Charles Darwin University. Work-experience placement project, November 2026.

## License

[MIT](LICENSE)
