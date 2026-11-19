import * as React from "react";
import * as ReactDOM from "react-dom";
import { Theme } from "@fluentui/react-components";
import { IEventFeedItem } from "@spteck/react-controls-v2";
import { ICalendarEventWithSource } from "../../models/ICalendarEventWithSource";

/**
 * EventHoverLayer
 *
 * Wraps the EventFeed control and shows a details card when the pointer
 * rests on an event card (or when an event card receives keyboard focus).
 *
 * The EventFeed control has no per-card hover hook, so the layer finds the
 * hovered card by walking up from the event target to the nearest card
 * element that contains exactly one event title.
 *
 * For SharePoint list calendars the layer also loads the extra submission
 * fields (category, campus, audience, registration, ...) on first hover and
 * caches them per event.
 */

export interface IEventHoverLayerProps {
  items: IEventFeedItem[];
  rawEvents: Map<string, ICalendarEventWithSource>;
  theme: Theme;
  enabled: boolean;
  children: React.ReactNode;
}

interface IExtraDetails {
  category?: string;
  deliveryMode?: string;
  campus?: string;
  audience?: string;
  organisingUnit?: string;
  organiser?: string;
  cost?: string;
  accessibility?: string;
  registrationRequired?: boolean;
  registrationCloses?: string;
  registrationUrl?: string;
  onlineMeetingUrl?: string;
  moreInfoUrl?: string;
}

interface IActiveCard {
  item: IEventFeedItem;
  rect: DOMRect;
}

const SHOW_DELAY_MS = 350;
const HIDE_DELAY_MS = 200;
const CARD_WIDTH = 340;
const GAP = 12;

const extraCache = new Map<string, IExtraDetails | null>();

function clean(value: unknown): string | undefined {
  if (value === undefined || value === null) return undefined;
  const text = String(value).trim();
  return text.length > 0 ? text : undefined;
}

function urlOf(value: unknown): string | undefined {
  if (value && typeof value === "object" && "Url" in (value as Record<string, unknown>)) {
    return clean((value as Record<string, unknown>).Url);
  }
  return undefined;
}

async function loadSharePointDetails(raw: ICalendarEventWithSource): Promise<IExtraDetails | null> {
  if (!raw._siteUrl || !raw._listId || !raw.id) return null;
  const key = `${raw._listId}|${raw.id}`;
  if (extraCache.has(key)) return extraCache.get(key) ?? null;

  const base = `${raw._siteUrl}/_api/web/lists(guid'${raw._listId}')/items(${raw.id})`;
  const init: RequestInit = {
    credentials: "same-origin",
    headers: { Accept: "application/json;odata=nometadata" },
  };
  try {
    const [textRes, itemRes] = await Promise.all([
      fetch(`${base}/FieldValuesAsText`, init),
      fetch(base, init),
    ]);
    if (!textRes.ok || !itemRes.ok) {
      extraCache.set(key, null);
      return null;
    }
    const text = (await textRes.json()) as Record<string, unknown>;
    const item = (await itemRes.json()) as Record<string, unknown>;
    const details: IExtraDetails = {
      category: clean(text.Category),
      deliveryMode: clean(text.DeliveryMode),
      campus: clean(text.Campus),
      audience: clean(text.Audience),
      organisingUnit: clean(text.OrganisingUnit),
      organiser: clean(text.OrganiserContact),
      cost: clean(text.EventCost),
      accessibility: clean(text.AccessibilityInfo),
      registrationRequired: item.RegistrationRequired === true,
      registrationCloses: clean(item.RegistrationCloses),
      registrationUrl: urlOf(item.RegistrationLink),
      onlineMeetingUrl: urlOf(item.OnlineMeetingLink),
      moreInfoUrl: urlOf(item.MoreInfo),
    };
    extraCache.set(key, details);
    return details;
  } catch {
    extraCache.set(key, null);
    return null;
  }
}

function formatWhen(item: IEventFeedItem): string {
  const start = new Date(item.startDate);
  const end = item.endDate ? new Date(item.endDate) : undefined;
  const day = new Intl.DateTimeFormat(undefined, { weekday: "short", day: "numeric", month: "short", year: "numeric" });
  const time = new Intl.DateTimeFormat(undefined, { hour: "numeric", minute: "2-digit" });
  if (isNaN(start.getTime())) return "";
  if (item.isAllDay) {
    if (!end || start.toDateString() === end.toDateString()) return `${day.format(start)} · All day`;
    return `${day.format(start)} – ${day.format(end)} · All day`;
  }
  if (!end || isNaN(end.getTime())) return `${day.format(start)}, ${time.format(start)}`;
  if (start.toDateString() === end.toDateString()) {
    return `${day.format(start)}, ${time.format(start)} – ${time.format(end)}`;
  }
  return `${day.format(start)}, ${time.format(start)} – ${day.format(end)}, ${time.format(end)}`;
}

function formatDateOnly(iso: string | undefined): string | undefined {
  if (!iso) return undefined;
  const d = new Date(iso);
  if (isNaN(d.getTime())) return undefined;
  return new Intl.DateTimeFormat(undefined, { weekday: "short", day: "numeric", month: "short" }).format(d);
}

/** Finds the card element under `target` and the event it shows. */
function findCard(
  target: HTMLElement,
  container: HTMLElement,
  items: IEventFeedItem[]
): { element: HTMLElement; item: IEventFeedItem } | undefined {
  let fallback: { element: HTMLElement; item: IEventFeedItem } | undefined;
  if (!container.contains(target)) return undefined;
  let el: HTMLElement | null = target;
  while (el && el !== container) {
    const text = el.textContent ?? "";
    const matches = items.filter((i) => !!i.title && text.indexOf(i.title) !== -1);
    if (matches.length === 1) {
      const isCard = el.getAttribute("role") === "button" && el.tagName !== "BUTTON";
      if (isCard) return { element: el, item: matches[0] };
      fallback = { element: el, item: matches[0] };
    } else if (matches.length > 1) {
      break;
    }
    el = el.parentElement;
  }
  return fallback;
}

export const EventHoverLayer: React.FC<IEventHoverLayerProps> = ({
  items,
  rawEvents,
  theme,
  enabled,
  children,
}) => {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const popupRef = React.useRef<HTMLDivElement>(null);
  const showTimer = React.useRef<number | undefined>(undefined);
  const hideTimer = React.useRef<number | undefined>(undefined);
  const pendingId = React.useRef<string | undefined>(undefined);
  const [active, setActive] = React.useState<IActiveCard | undefined>(undefined);
  const [extra, setExtra] = React.useState<IExtraDetails | null | undefined>(undefined);
  const [position, setPosition] = React.useState<{ top: number; left: number } | undefined>(undefined);

  const clearTimers = (): void => {
    window.clearTimeout(showTimer.current);
    window.clearTimeout(hideTimer.current);
  };

  const hide = React.useCallback((): void => {
    clearTimers();
    pendingId.current = undefined;
    setActive(undefined);
    setPosition(undefined);
  }, []);

  const scheduleHide = (): void => {
    window.clearTimeout(showTimer.current);
    pendingId.current = undefined;
    window.clearTimeout(hideTimer.current);
    hideTimer.current = window.setTimeout(hide, HIDE_DELAY_MS);
  };

  const open = (element: HTMLElement, item: IEventFeedItem, delay: number): void => {
    window.clearTimeout(hideTimer.current);
    if (active?.item.id === item.id || pendingId.current === item.id) return;
    window.clearTimeout(showTimer.current);
    pendingId.current = item.id;
    showTimer.current = window.setTimeout(() => {
      pendingId.current = undefined;
      setPosition(undefined);
      setActive({ item, rect: element.getBoundingClientRect() });
    }, delay);
  };

  const onMouseOver = (ev: React.MouseEvent<HTMLDivElement>): void => {
    if (!enabled || !containerRef.current) return;
    // Events from the portal-rendered card bubble here through React; keep it open.
    if (popupRef.current && popupRef.current.contains(ev.target as Node)) {
      window.clearTimeout(hideTimer.current);
      return;
    }
    const found = findCard(ev.target as HTMLElement, containerRef.current, items);
    if (found) open(found.element, found.item, SHOW_DELAY_MS);
    else scheduleHide();
  };

  const onFocus = (ev: React.FocusEvent<HTMLDivElement>): void => {
    if (!enabled || !containerRef.current) return;
    if (popupRef.current && popupRef.current.contains(ev.target as Node)) return;
    const found = findCard(ev.target as HTMLElement, containerRef.current, items);
    if (found) open(found.element, found.item, 0);
  };

  // Load SharePoint submission fields for the active event.
  React.useEffect(() => {
    if (!active) return;
    const raw = rawEvents.get(active.item.id);
    setExtra(undefined);
    if (!raw || !raw._siteUrl) {
      setExtra(null);
      return;
    }
    let cancelled = false;
    void loadSharePointDetails(raw).then((d) => {
      if (!cancelled) setExtra(d);
    });
    return () => {
      cancelled = true;
    };
  }, [active, rawEvents]);

  // Place the card beside the event once its height is known.
  React.useLayoutEffect(() => {
    if (!active || !popupRef.current) return;
    const { rect } = active;
    const height = popupRef.current.offsetHeight;
    const vw = window.innerWidth;
    const vh = window.innerHeight;
    let left = rect.right + GAP;
    let top = rect.top;
    if (left + CARD_WIDTH > vw - 8) left = rect.left - CARD_WIDTH - GAP;
    if (left < 8) {
      left = Math.min(Math.max(rect.left, 8), vw - CARD_WIDTH - 8);
      top = rect.bottom + GAP;
      if (top + height > vh - 8) top = rect.top - height - GAP;
    }
    top = Math.min(Math.max(top, 8), Math.max(8, vh - height - 8));
    setPosition({ top, left });
  }, [active, extra]);

  // Hide on Escape and when the page scrolls (the card would drift).
  React.useEffect(() => {
    if (!active) return;
    const onKey = (e: KeyboardEvent): void => {
      if (e.key === "Escape") hide();
    };
    const onScroll = (e: Event): void => {
      if (popupRef.current && e.target instanceof Node && popupRef.current.contains(e.target)) return;
      hide();
    };
    document.addEventListener("keydown", onKey);
    window.addEventListener("scroll", onScroll, true);
    return () => {
      document.removeEventListener("keydown", onKey);
      window.removeEventListener("scroll", onScroll, true);
    };
  }, [active, hide]);

  React.useEffect(() => () => clearTimers(), []);

  if (!enabled) return <>{children}</>;

  const item = active?.item;
  const raw = item ? rawEvents.get(item.id) : undefined;
  const t = theme;
  const category = extra?.category ?? item?.category;
  const organiser = extra?.organiser ?? item?.organizer;
  const where = [item?.location, extra?.campus && extra.campus !== item?.location ? extra.campus : undefined]
    .filter(Boolean)
    .join(" · ");

  const rows: Array<[string, string | undefined]> = item
    ? [
        ["When", formatWhen(item)],
        ["Where", where || undefined],
        ["Format", extra?.deliveryMode ?? (item.isVirtual ? "Online" : undefined)],
        ["For", extra?.audience],
        ["Organiser", [organiser, extra?.organisingUnit].filter(Boolean).join(", ") || undefined],
        ["Cost", extra ? extra.cost ?? "Free" : undefined],
        ["Access", extra?.accessibility],
        [
          "Registration",
          extra?.registrationRequired
            ? `Required${extra.registrationCloses ? `, closes ${formatDateOnly(extra.registrationCloses)}` : ""}`
            : undefined,
        ],
      ]
    : [];

  const links: Array<[string, string | undefined]> = item
    ? [
        ["Register", extra?.registrationUrl],
        ["Join online", extra?.onlineMeetingUrl ?? item.joinUrl],
        ["More information", extra?.moreInfoUrl],
        ["Open event", item.linkUrl],
      ]
    : [];

  const linkStyle: React.CSSProperties = {
    color: t.colorBrandForegroundLink,
    fontWeight: 600,
    textDecoration: "none",
    marginRight: 14,
    whiteSpace: "nowrap",
  };

  const popup =
    item &&
    ReactDOM.createPortal(
      <div
        ref={popupRef}
        role="dialog"
        aria-label={`Event details: ${item.title}`}
        onMouseEnter={() => window.clearTimeout(hideTimer.current)}
        onMouseLeave={scheduleHide}
        style={{
          position: "fixed",
          top: position?.top ?? -9999,
          left: position?.left ?? -9999,
          width: CARD_WIDTH,
          maxHeight: "80vh",
          overflowY: "auto",
          zIndex: 1000000,
          background: t.colorNeutralBackground1,
          color: t.colorNeutralForeground1,
          borderRadius: 8,
          boxShadow: t.shadow28,
          border: `1px solid ${t.colorNeutralStroke2}`,
          fontFamily: t.fontFamilyBase,
          fontSize: 13,
          lineHeight: "18px",
          visibility: position ? "visible" : "hidden",
        }}
      >
        {item.imageUrl && (
          <div
            style={{
              height: 120,
              backgroundImage: `url("${item.imageUrl}")`,
              backgroundSize: "cover",
              backgroundPosition: "center",
              borderTopLeftRadius: 8,
              borderTopRightRadius: 8,
            }}
          />
        )}
        <div style={{ padding: "12px 16px 14px" }}>
          {category && (
            <span
              style={{
                display: "inline-block",
                marginBottom: 6,
                padding: "2px 8px",
                borderRadius: 10,
                fontSize: 11,
                fontWeight: 600,
                textTransform: "uppercase",
                letterSpacing: 0.4,
                background: t.colorBrandBackground2,
                color: t.colorBrandForeground2,
              }}
            >
              {category}
            </span>
          )}
          <div style={{ fontSize: 16, lineHeight: "22px", fontWeight: 600, marginBottom: 8 }}>{item.title}</div>
          <dl style={{ margin: 0, display: "grid", gridTemplateColumns: "88px 1fr", rowGap: 4, columnGap: 8 }}>
            {rows
              .filter(([, value]) => !!value)
              .map(([label, value]) => (
                <React.Fragment key={label}>
                  <dt style={{ color: t.colorNeutralForeground3, margin: 0 }}>{label}</dt>
                  <dd style={{ margin: 0 }}>{value}</dd>
                </React.Fragment>
              ))}
          </dl>
          {item.description && (
            <p style={{ margin: "10px 0 0", color: t.colorNeutralForeground2, whiteSpace: "pre-line" }}>
              {item.description}
            </p>
          )}
          {extra === undefined && raw?._siteUrl && (
            <div style={{ marginTop: 8, color: t.colorNeutralForeground3, fontSize: 12 }}>Loading more details…</div>
          )}
          {links.some(([, url]) => !!url) && (
            <div style={{ marginTop: 12, paddingTop: 10, borderTop: `1px solid ${t.colorNeutralStroke2}` }}>
              {links
                .filter(([, url]) => !!url)
                .map(([label, url]) => (
                  <a key={label} href={url} target="_blank" rel="noopener noreferrer" style={linkStyle}>
                    {label}
                  </a>
                ))}
            </div>
          )}
        </div>
      </div>,
      document.body
    );

  return (
    <div
      ref={containerRef}
      onMouseOver={onMouseOver}
      onMouseLeave={scheduleHide}
      onFocus={onFocus}
      onBlur={(ev) => {
        const next = ev.relatedTarget as Node | null;
        if (next && (containerRef.current?.contains(next) || popupRef.current?.contains(next))) return;
        scheduleHide();
      }}
    >
      {children}
      {popup}
    </div>
  );
};

export default EventHoverLayer;
