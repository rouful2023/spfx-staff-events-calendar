# Setting up the staff calendar in SharePoint

These are the steps used to build the test site. A site owner can follow them on any team or communication site.

## 1. Theme

Apply a custom theme with a CDU red primary colour (`#A71930`) and set the site header background to **Strong** (Settings → Change the look).

## 2. Events list

Create an **Events** list called *Staff Events Calendar*. Rename these built-in columns:

| Built-in column | New display name |
|---|---|
| Title | Event title (maximum length 80) |
| EventDate / EndDate | Start date and time / End date and time |
| Description | Short description (required) |
| Location | Room or location |
| BannerUrl | Image |

Set the **Category** choices to:

- Academic and key dates
- Student events
- Staff events and meetings
- Professional development
- Social and wellbeing
- Operational
- Community and external

Make Category required.

Add these columns:

| Internal name | Display name | Type | Required |
|---|---|---|---|
| DeliveryMode | Delivery mode | Choice: In person, Online, Hybrid | Yes |
| Campus | Campus | Choice (fill-in allowed) | Yes |
| Audience | Audience | Multi-choice: All staff, Academic staff, Professional staff, Students, External | Yes |
| OrganisingUnit | Organising unit | Choice (fill-in allowed) | Yes |
| OrganiserContact | Organiser or contact | Person | Yes |
| RegistrationRequired | Registration required | Yes/No | No |
| RegistrationLink | Registration link | Hyperlink | No |
| RegistrationCloses | Registration closes | Date only | No |
| OnlineMeetingLink | Online meeting link | Hyperlink | No |
| EventCost | Cost | Single line | No |
| AccessibilityInfo | Accessibility information | Multiple lines | No |
| MoreInfo | More information | Hyperlink | No |
| RepeatsNote | Repeats | Single line | No |

Hide the unused built-in columns from the new and edit forms: Geolocation, ParticipantsPicker, Facilities, FreeBusy, Overbook and BannerImageUrl.

The hover card reads columns by these internal names. If you use different names, update `loadSharePointDetails` in `src/components/EventHoverCard/EventHoverCard.tsx`.

## 3. Approval

Go to List settings → Versioning settings and set:

- **Require content approval for submitted items:** Yes
- **Create a version each time you edit an item:** Yes
- **Who should see draft items:** Only users who can approve items, and the author

SharePoint then adds the **My submissions** and **Approve/reject Items** views.

## 4. Views

- **Calendar:** the default month view.
- **Upcoming (agenda):** filter End date ≥ [Today], sort by Start date. Show Category, Delivery mode, Campus, Location, Audience, Organising unit and Approval status.
- **By category:** the same filter, grouped by Category.

## 5. Page

Create a page and add:

- a banner, with a gradient section background
- **Events Feed** (this web part) in the agenda layout, pointed at the list (Data source → select calendar)
- in the right column: a **Button** ("Submit an event", linking to the list's `NewForm.aspx`), a **Text** web part with links to the views, **World clock** and **Countdown timer**
- a second section with Events Feed in the carousel, grid (drag to reorder) and compact list layouts

Until the `.sppkg` is deployed to the app catalog, the Events Feed web part only loads while `npm run start` is running and the page is opened with `?debugManifestsFile=...&loadSPFX=true`.
