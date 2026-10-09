"""Build one static demo page per business.

Full run:   python build.py --checked columbus_targets_v3_checked.xlsx
Samples:    python build.py --samples "Business A" "Business B"   (writes pages only, no xlsx)
"""
import argparse, html, json, re, shutil
from pathlib import Path
import pandas as pd

HERE = Path(__file__).parent
DL = Path(r"C:\Users\misha\Downloads")
OUTSCRAPER = DL / "Outscraper-20261006150726s1c91.xlsx"
MY_EMAIL = "mishabichashvili1998@gmail.com"
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
E = html.escape
MISSED_CALL = ("HVAC contractor", "Plumber", "handyman")


# ---------- listing data helpers ----------
SLUG_DROP = {"llc", "inc", "heating", "services", "co", "the", "and"}


def slug_words(name):
    words = []
    for tok in name.lower().replace("'", "").replace("’", "").split():
        w = re.sub(r"-+", "-", re.sub(r"[^a-z0-9]+", "-", tok)).strip("-")
        if w and w not in SLUG_DROP:
            words.append(w)
    return words or [re.sub(r"[^a-z0-9]+", "-", name.lower()).strip("-")]


def make_slugs(names):
    """First two meaningful words; colliding groups get a third word, then a numeric suffix."""
    words = {n: slug_words(n) for n in names}
    out = {n: "-".join(w[:2]) for n, w in words.items()}
    for k in (3, 4):
        counts = {}
        for v in out.values():
            counts[v] = counts.get(v, 0) + 1
        for n in names:
            if counts[out[n]] > 1:
                out[n] = "-".join(words[n][:k])
    final, seen = {}, {}
    for n in names:
        sl = out[n]
        seen[sl] = seen.get(sl, 0) + 1
        final[n] = sl if seen[sl] == 1 else f"{sl}-{seen[sl]}"
    return final


def loads(s):
    try:
        return json.loads(s) if isinstance(s, str) else {}
    except Exception:
        return {}


def to_min(t, mer):
    h, _, m = t.partition(":")
    h = int(h) % 12 + (12 if mer.upper() == "PM" else 0)
    return h * 60 + int(m or 0)


def parse_range(s):
    """'7:30AM-4:30PM' -> (450, 990); None if closed/unparseable. '24' -> (0, 1440)."""
    if not s or "closed" in s.lower():
        return None
    if "24" in s and "hour" in s.lower():
        return (0, 1440)
    m = re.match(r"\s*([\d:]+)\s*([AP]M)?\s*[-\u2013]\s*([\d:]+)\s*([AP]M)", s, re.I)
    if not m:
        return None
    a, am, b, bm = m.groups()
    return (to_min(a, am or bm), to_min(b, bm))


def fmt_min(x):
    h, m = divmod(x % 1440, 60)
    return f"{(h % 12) or 12}:{m:02d} {'AM' if h < 12 else 'PM'}".replace(":00", "")


def day_hours(hours, day):
    v = hours.get(day)
    if isinstance(v, list):
        v = v[0] if v else None
    return parse_range(v)


def hours_summary(hours):
    """Collapse to e.g. 'Mon-Fri 8 AM to 6 PM'. None if no hours."""
    runs, prev = [], None
    for d in DAYS:
        r = day_hours(hours, d)
        if prev and prev[0] == r:
            prev[1].append(d)
        else:
            prev = [r, [d]]
            runs.append(prev)
    parts = []
    for r, ds in runs:
        if r is None:
            continue
        label = ds[0][:3] if len(ds) == 1 else f"{ds[0][:3]}-{ds[-1][:3]}"
        if len(ds) == 7:
            label = "Every day"
        parts.append(f"{label} {fmt_min(r[0])} to {fmt_min(r[1])}" if r != (0, 1440) else f"{label} 24 hours")
    return "; ".join(parts) or None


def next_open(hours, after="Tuesday"):
    """First open day after `after` -> (day name, open minute)."""
    i = DAYS.index(after)
    for k in range(1, 8):
        d = DAYS[(i + k) % 7]
        r = day_hours(hours, d)
        if r:
            return d, r[0]
    return "Wednesday", 9 * 60


def city_of(address):
    parts = [p.strip() for p in str(address).split(",")]
    return parts[-2] if len(parts) >= 3 else "Columbus"


def offerings(about):
    out = []
    for grp, vals in about.items():
        if grp.lower().startswith("offering") or grp.lower().startswith("service"):
            out += [k for k, v in vals.items() if v and k.lower() not in ("onsite services",)]
    return out


def subtype_list(subtypes):
    return [s.strip() for s in str(subtypes).split(",") if s.strip()]


SUFFIX = re.compile(r"\s+(contractor|service|services|designer)$", re.I)
GENERIC_TYPES = {"contractor", "service", "services", "company", "business", "general"}
NICE = {"landscape": "landscaping", "landscaper": "landscaping", "gardener": "gardening", "tree": "tree work", "handyman": "handyman work",
        "plumber": "plumbing", "roofer": "roofing", "painter": "painting", "electrician": "electrical work"}


def natural_list(items):
    return items[0] if len(items) == 1 else ", ".join(items[:-1]) + " and " + items[-1]


def nice_subtypes(subs, drop=()):
    """['Landscaper', 'Masonry contractor', 'Tree service'] -> 'landscaping, masonry and tree work'."""
    out = []
    for s_ in subs:
        t = SUFFIX.sub("", s_.strip().lower())
        t = NICE.get(t, t)
        if not t or t in GENERIC_TYPES or "handy" in t:
            continue
        if t in drop or t in out or any(t[:6] == o[:6] for o in out):
            continue
        out.append(t)
    return natural_list(out) if out else ""


# ---------- per-trade content ----------
def build_content(r):
    name, trade = r["name"], r["trade"]
    hours, about = loads(r.get("working_hours")), loads(r.get("about"))
    subs = subtype_list(r.get("subtypes"))
    city = city_of(r["address"])
    first_day, first_open = next_open(hours)
    slot = fmt_min(first_open + 60 if first_open + 60 < 22 * 60 else first_open)
    slot_day = first_day
    tue = day_hours(hours, "Tuesday")
    after_hours = bool(tue) and tue[1] <= 20 * 60
    sat = day_hours(hours, "Saturday")
    lower = f"{name} {' '.join(subs)}".lower()

    if sat and sat != (0, 1440):
        sat_ans = f"Yes, we are open Saturday, {fmt_min(sat[0])} to {fmt_min(sat[1])}. Share your address and we will find you a slot."
    elif sat:
        sat_ans = "Yes, we are open Saturday. Share your address and we will find you a slot."
    else:
        sd, so = next_open(hours, "Saturday")
        sat_ans = f"We are closed on Saturday. Our next opening is {sd} at {fmt_min(so)}, and I can book that for you now."
    area_ans = f"We are based in {city}, Ohio. Send us your address and we will confirm we can get to you."
    soon_ans = f"Our next available slot is {slot_day} at {slot}. Reply with your address and I will hold it."

    thread, qa = [], []
    ai_open = f"Hi, this is {name}. Sorry we missed your call."
    when = "Tue 9:12 PM" if after_hours else "Tue 1:20 PM"
    why = "We are closed for the night, but I can help right now." if after_hours else "The team is out on a job, but I can help right now."

    if trade in ("HVAC contractor", "Plumber", "handyman"):
        if trade == "HVAC contractor":
            heat = "heat" in lower or "furnace" in lower
            ask = ("My furnace is not kicking on and the house is getting cold." if heat
                   else "My AC is running but only blowing warm air and the house will not cool down.")
            job = "Furnace not heating" if heat else "AC blowing warm air"
            off = offerings(about)
            names = [re.sub(r"\s+services?$", "", o, flags=re.I).lower() for o in off[:2]]
            ans = ("Yes, we handle " + (" and ".join(names) if names else "repairs and installs")
                   + " for home heating and cooling. Is the thermostat showing anything?")
            tq = ("Do you install new systems too?", f"Our listing shows {' and '.join(names)}. Tell us what you have now and what you are after, and we will line up the right visit." if off
                  else "Yes. We work on home heating and cooling systems. Tell us what you have now and what you are after, and we will line up the right visit.")
            follow = "The thermostat is on but nothing happens. Tomorrow morning would be best."
        elif trade == "Plumber":
            ask = "I have a leak under my kitchen sink and the cabinet floor is getting wet."
            job = "Leak under kitchen sink"
            ans = "Thanks, that is something we can look at. Is water still running right now?"
            tq = ("Do you handle leaks and clogged drains?", "Yes. Leaks, clogs and fixture repairs are common calls for us. Text a short description and a photo if you can.")
            follow = "It is slow, I shut the valve off. Tomorrow morning would be best."
        else:
            ask = "I have a few small repairs around the house that I keep putting off."
            job = "Faucet repair and sticking door"
            kinds = nice_subtypes(subs, drop=("handyman work",))
            ans = "Happy to help with that. What is on the list?"
            tq = ("What kind of jobs do you take on?", "We take on repairs and small home projects"
                  + (f", and our listing also shows {kinds}" if kinds else "")
                  + ". Send us your list and we will tell you what fits.")
            follow = "A leaky faucet and a sticking door. Tomorrow morning would be best."
        thread = [("sys", f"Missed call from (614) 555-0142 \u00b7 {when}"),
                  ("ai", f"{ai_open} {why} What is going on?"),
                  ("cust", ask), ("ai", ans), ("cust", follow),
                  ("ai", f"Our next opening is {slot_day} at {slot}. Want me to hold it? I just need your address."),
                  ("cust", "Yes please. 1420 Maple Ave, Columbus 43214."),
                  ("ai", f"You are booked for {slot_day} at {slot} at 1420 Maple Ave. We will text when we are on the way.")]
        qa = [("What areas do you cover?", area_ans), ("Can someone come Saturday?", sat_ans),
              ("How soon can you get here?", soon_ans), tq]

    elif trade in ("Landscaper", "Lawn care service"):
        if trade == "Landscaper":
            ask = "Hi, I would like a quote for a fall cleanup and fresh mulch in my beds."
            job = "Fall cleanup and mulch estimate"
            size = "About a quarter acre. One time cleanup now, maybe weekly mowing in spring. Can you also trim the hedges along the fence?"
            extra = "I will add the hedges to the estimate so the team can look at them in person."
        else:
            ask = "Hi, I would like a quote for weekly mowing starting in spring."
            job = "Weekly mowing estimate"
            size = "About a quarter acre, front and back. Weekly mowing. Does that usually include edging along the driveway?"
            extra = "I will note edging on the estimate and the team will go over what is included when they visit."
        kinds = nice_subtypes(subs)
        scope_q = "Do you do more than one kind of yard work?"
        scope_a = (f"Our listing shows {kinds}." if " and " in kinds
                   else "Yard and property care is our core work, so cleanups, beds and regular service all fit.")
        thread = [("sys", f"Text received \u00b7 Tue {'6:48 PM' if after_hours else '1:20 PM'}"),
                  ("cust", ask),
                  ("ai", f"Thanks for texting {name}. I can set up an estimate visit. What is the property address?"),
                  ("cust", "1420 Maple Ave, Columbus 43214."),
                  ("ai", "Thanks. Roughly how big is the yard, and is this a one time job or regular service?"),
                  ("cust", size),
                  ("ai", f"That helps. {extra} I can send someone out {slot_day} at {slot} for an estimate. Does that work?"),
                  ("cust", f"{slot_day} works."),
                  ("ai", f"Booked. An estimate visit is set for {slot_day} at {slot} at 1420 Maple Ave. We will text before arriving.")]
        qa = [("What areas do you cover?", area_ans), ("Can someone come Saturday?", sat_ans),
              ("How soon can you come out for an estimate?", f"Our next estimate slot is {slot_day} at {slot}. Reply with your address and I will hold it."),
              (scope_q, scope_a + " Tell us what you have in mind and we will say what fits.")]

    elif trade == "roofing contractor":
        thread = [("sys", f"Text received \u00b7 Tue {'9:12 PM' if after_hours else '1:20 PM'}"),
                  ("cust", "That storm was rough. I am worried my roof got damaged."),
                  ("ai", f"Thanks for texting {name}. I am sorry to hear that. Do you see any active leak or water inside right now?"),
                  ("cust", "No leak, but I see a few shingles in the yard."),
                  ("ai", "Good that it is not leaking. A roof inspection is the next step. Do you need to know anything first?"),
                  ("cust", "Will my insurance company want a report from you?"),
                  ("ai", f"Insurers usually want photos and notes on the damage. The team can walk you through that at the inspection, and I will flag it so they come prepared. I can book {slot_day} at {slot}. Does that work?"),
                  ("cust", "Yes, book it. 1420 Maple Ave."),
                  ("ai", f"Booked. A roof inspection is set for {slot_day} at {slot} at 1420 Maple Ave. If a leak starts, text us here.")]
        qa = [("What areas do you cover?", area_ans), ("Can someone come Saturday?", sat_ans),
              ("How soon can you inspect my roof?", f"Our next inspection slot is {slot_day} at {slot}. Reply with your address and I will hold it."),
              ("What should I do after a storm?", "Take photos from the ground if it is safe, and text us what you see, such as missing shingles or a leak. We will book an inspection and the team can talk through next steps there.")]

    else:  # house cleaning service
        thread = [("sys", f"Text received \u00b7 Tue {'9:12 PM' if after_hours else '1:20 PM'}"),
                  ("cust", "Hi, I am looking for a first cleaning for my house."),
                  ("ai", f"Thanks for texting {name}. Happy to help. About how big is the home, and how often would you like us to come?"),
                  ("cust", "Three bedrooms, two baths. Every two weeks. We have a dog."),
                  ("ai", "Got it. A dog is no problem. Do you have questions before we book?"),
                  ("cust", "Do you bring your own supplies?"),
                  ("ai", f"I will confirm the supplies details with the team and text you before your first visit. I can hold a recurring slot every other {slot_day} at {slot}. Does that work?"),
                  ("cust", f"Yes, every other {slot_day} at {slot}. The address is 1420 Maple Ave."),
                  ("ai", f"Booked. Your first cleaning is {slot_day} at {slot} at 1420 Maple Ave, then every two weeks.")]
        qa = [("What areas do you cover?", area_ans), ("Can someone come Saturday?", sat_ans),
              ("How soon can you start?", f"Our next opening is {slot_day} at {slot}. Reply with your address and home size and I will hold it."),
              ("Do you offer recurring cleanings?", "Yes. Tell us the home size and how often you would like us, such as weekly, every two weeks or monthly, and we will set up a recurring slot.")]
    job = locals().get("job") or {"roofing contractor": "Storm damage roof inspection",
                                  "house cleaning service": "Recurring cleaning, 3 bed / 2 bath"}.get(trade, "New job")
    notify = f"New job booked \u2014 {job}, 1420 Maple Ave, {slot_day} {slot}. Customer: (614) 555-0142"
    return thread, qa, hours_summary(hours), city, notify


# ---------- rendering ----------
def render(r):
    name = E(r["name"])
    thread, qa, hsum, city, notify = build_content(r)
    bub = []
    for kind, text in thread:
        if kind == "sys":
            bub.append(f'<p class="sys">{E(text)}</p>')
        else:
            label = '<span class="who">AI receptionist</span>' if kind == "ai" else ""
            bub.append(f'<p class="msg {kind}">{label}{E(text)}</p>')
    bits = []
    if pd.notna(r.get("rating")) and pd.notna(r.get("reviews")):
        bits.append(f"{float(r['rating']):.1f} stars from {int(r['reviews'])} Google reviews")
    if hsum:
        bits.append(f"Hours: {hsum}")
    bits.append(f"Based in {E(city)}, OH")
    buttons = "".join(f'<button type="button" aria-expanded="false" aria-controls="a{i}">{E(q)}</button>'
                      f'<p class="ans" id="a{i}" hidden>{E(a)}</p>' for i, (q, a) in enumerate(qa))
    subject = f"demo for {r['name']}".replace(" ", "%20").replace("&", "%26")
    return f"""<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="robots" content="noindex,nofollow">
<title>AI receptionist demo for {name}</title>
<link rel="stylesheet" href="../style.css">
</head>
<body>
<main>
<h1>What a {"missed call" if r["trade"] in MISSED_CALL else "quote request"} to {name} looks like with an AI receptionist</h1>
<p class="facts">{E(' \u00b7 '.join(bits))}</p>
<div class="phone" aria-label="Example text conversation">
<div class="bar">{name}</div>
<div class="thread">
{chr(10).join(bub)}
</div>
</div>
<p class="note">Example conversation. The customer and address are made up.</p>
<div class="notify"><span class="nlabel">Text to the owner</span>{E(notify)}</div>
<h2>Ask it something a customer would ask</h2>
<div class="qa">{buttons}</div>
<section class="runs">
<h2>This runs on your existing phone number. Nothing to install.</h2>
<ul>
<li>Missed calls get a text back in 30 seconds</li>
<li>Answers questions from your Google listing</li>
<li>Books jobs into your calendar</li>
</ul>
</section>
<a class="cta" href="mailto:{MY_EMAIL}?subject={subject}">See it work for {name}</a>
</main>
<footer>Demo built from public Google listing data. Not affiliated with {name}.</footer>
<script>
document.querySelectorAll(".qa button").forEach(function (b) {{
  b.addEventListener("click", function () {{
    var a = document.getElementById(b.getAttribute("aria-controls"));
    var open = a.hidden;
    document.querySelectorAll(".qa .ans").forEach(function (x) {{ x.hidden = true; }});
    document.querySelectorAll(".qa button").forEach(function (x) {{ x.setAttribute("aria-expanded", "false"); }});
    a.hidden = !open;
    b.setAttribute("aria-expanded", String(open));
  }});
}});
</script>
</body>
</html>
"""


CSS = """:root{--ink:#1c2430;--mute:#5b6675;--bg:#f4f6f8;--card:#fff;--accent:#1f5eff;--line:#dde2e8}
*{box-sizing:border-box}
body{margin:0;background:var(--bg);color:var(--ink);font:16px/1.5 system-ui,-apple-system,"Segoe UI",Roboto,Arial,sans-serif}
main{max-width:560px;margin:0 auto;padding:24px 16px 8px}
h1{font-size:1.5rem;line-height:1.25;margin:0 0 8px}
h2{font-size:1.1rem;margin:28px 0 10px}
.facts{color:var(--mute);font-size:.9rem;margin:0 0 20px}
.phone{background:var(--card);border:2px solid var(--ink);border-radius:28px;overflow:hidden;box-shadow:0 6px 20px rgba(0,0,0,.08)}
.bar{background:var(--ink);color:#fff;text-align:center;padding:12px;font-weight:600;font-size:.95rem}
.thread{padding:14px 12px;display:flex;flex-direction:column;gap:8px}
.msg{margin:0;padding:9px 13px;border-radius:18px;max-width:84%;font-size:.95rem}
.msg.ai{background:#e8edf3;align-self:flex-start;border-bottom-left-radius:5px}
.msg.cust{background:var(--accent);color:#fff;align-self:flex-end;border-bottom-right-radius:5px}
.who{display:block;font-size:.7rem;font-weight:700;color:var(--mute);margin-bottom:2px;text-transform:uppercase;letter-spacing:.04em}
.sys{margin:4px 0;text-align:center;font-size:.78rem;color:var(--mute)}
.notify{margin:14px 0 0;padding:12px 14px;background:var(--card);border:1px solid var(--line);border-left:4px solid var(--accent);border-radius:10px;font-size:.9rem}
.nlabel{display:block;font-size:.7rem;font-weight:700;color:var(--mute);text-transform:uppercase;letter-spacing:.04em;margin-bottom:2px}
.note{font-size:.8rem;color:var(--mute);margin:8px 4px 0;text-align:center}
.qa button{display:block;width:100%;text-align:left;margin:0 0 8px;padding:13px 14px;background:var(--card);border:1px solid var(--line);border-radius:10px;font:inherit;font-weight:600;color:var(--ink);cursor:pointer}
.qa button[aria-expanded=true]{border-color:var(--accent)}
.qa button:focus-visible,.cta:focus-visible{outline:3px solid var(--accent);outline-offset:2px}
.ans{margin:-2px 0 12px;padding:12px 14px;background:#e8edf3;border-radius:10px}
.runs ul{margin:0;padding-left:20px}
.runs li{margin:6px 0}
.cta{display:block;margin:28px 0 8px;padding:15px;background:var(--accent);color:#fff;text-align:center;text-decoration:none;font-weight:700;border-radius:10px}
footer{max-width:560px;margin:0 auto;padding:16px;font-size:.75rem;color:var(--mute);text-align:center}
@media(prefers-color-scheme:dark){:root{--ink:#e8ecf1;--mute:#9aa6b5;--bg:#12161c;--card:#1b212a;--line:#2d3642}
.bar{background:#2d3642}.phone{border-color:#2d3642}.msg.ai,.ans{background:#2a323d}}
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--checked", help="path to columbus_targets_v3_checked.xlsx")
    ap.add_argument("--urls-out", help="write {name: /slug/} JSON here")
    ap.add_argument("--samples", nargs="*", help="business names; build only these, skip the xlsx")
    a = ap.parse_args()

    src = Path(a.checked) if a.checked else DL / "columbus_targets_v3.xlsx"
    df = pd.read_excel(src)
    if a.samples:
        rows = df[df["name"].isin(a.samples)]
    else:
        if "send" not in df.columns:
            raise SystemExit("No `send` column in the input; pass --checked columbus_targets_v3_checked.xlsx")
        rows = df[df["send"].astype(str).str.strip().str.upper() == "YES"]
    o = pd.read_excel(OUTSCRAPER).drop_duplicates("name")
    keep = ["name", "working_hours", "about", "description"]
    rows = rows.drop(columns=[c for c in keep[1:] if c in rows.columns]).merge(o[keep], on="name", how="left")

    for old in (HERE / "demos", HERE / "style.css"):
        if old.is_dir():
            shutil.rmtree(old)
        elif old.exists():
            old.unlink()
    for d in HERE.iterdir():  # drop pages from a previous run
        if d.is_dir() and (d / "index.html").exists() and not d.name.startswith("."):
            shutil.rmtree(d)
    out = HERE
    (out / "style.css").write_text(CSS, encoding="utf-8")
    slugs = make_slugs(list(rows["name"]))
    urls = {}
    for _, r in rows.iterrows():
        s = slugs[r["name"]]
        (out / s).mkdir()
        (out / s / "index.html").write_text(render(r), encoding="utf-8")
        urls[r["name"]] = f"/{s}/"
    (HERE / "index.html").write_text('<!doctype html>\n<html lang="en"><head><meta charset="utf-8">'
                                     '<meta name="robots" content="noindex,nofollow"><title>Columbus demos</title></head><body></body></html>\n',
                                     encoding="utf-8")
    (HERE / "robots.txt").write_text("User-agent: *\nDisallow: /\n", encoding="utf-8")
    print(f"built {len(urls)} pages")
    if a.urls_out:
        Path(a.urls_out).write_text(json.dumps(urls, indent=1), encoding="utf-8")


if __name__ == "__main__":
    main()
