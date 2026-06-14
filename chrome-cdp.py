#!/usr/bin/env python3
"""
chrome-cdp.py — drive phone Chrome over CDP + Outlook account creator.

Talks to Chrome's DevTools Protocol directly over a WebSocket (see cdp.py),
using only the pure-Python `websocket-client` — no Playwright/Node, no
chromedriver. That's what lets it run on raw Termux/Android too. Setup:
    pip install -r requirements.txt      # just websocket-client
Still needs `adb forward` to the devtools socket (handled automatically).

Commands:
    forward, info, tabs, open, goto, eval, text, shot, screen, fill, click
    signup [--brightdata]   Create Outlook account (and optional BrightData signup)

Global options:
    --serial IP:5555        adb serial
    --port 9222             local forwarding port
    --socket NAME           remote devtools socket
    --incognito             fresh Microsoft/Outlook session (clears their cookies
                            only; true incognito isn't possible over Android CDP).
                            Works before or after the subcommand.
"""

import argparse
import json
import os
import random
import re
import subprocess
import sys
import time
import urllib.request

import cdp

# ── Names ──────────────────────────────────────────────────────────────
FIRST_NAMES = [
    "Aarav", "Vivaan", "Aditya", "Arjun", "Sai", "Reyansh", "Ayaan", "Krishna",
    "Ishaan", "Shaurya", "Atharv", "Dhruv", "Kabir", "Rudra", "Parth", "Yash",
    "Dev", "Mohammed", "Rohan", "Aryan", "Harsh", "Om", "Veer", "Amit", "Deepak",
    "Rahul", "Vikram", "Siddharth", "Manoj", "Abhishek", "Yusuf", "Anil", "Sunil",
    "Raj", "Nitin", "Prakash", "Karan", "Varun", "Gaurav", "Shubham", "Nikhil",
    "Aniket", "Pranav", "Tanmay", "Sagar", "Tejas", "Sahil", "Tushar", "Laksh", "Avi"
]
LAST_NAMES = [
    "Sharma", "Verma", "Singh", "Gupta", "Patel", "Kumar", "Jain", "Agarwal",
    "Reddy", "Rao", "Nair", "Menon", "Iyer", "Mishra", "Pandey", "Tiwari",
    "Yadav", "Saxena", "Srivastava", "Bose", "Das", "Chakraborty", "Mukherjee",
    "Chatterjee", "Banerjee", "Sen", "Thakur", "Rajput", "Chauhan", "Rathore",
    "Joshi", "Kulkarni", "Deshmukh", "Pawar", "Gaikwad", "Shinde", "Patil",
    "Kamath", "Hegde", "Shetty", "Khan", "Syed", "Ansari", "Qureshi", "Mirza"
]

# ── Helpers ────────────────────────────────────────────────────────────
def gen_email(first: str, last: str) -> str:
    return f"{first.lower()}{last.lower()}{random.randint(0,9999):04d}@outlook.com"

def gen_password(first: str) -> str:
    base = first.lower()
    while len(base) < 5:
        base += random.choice("abcdefghijklmnopqrstuvwxyz")
    return f"{base}{random.randint(0,999):03d}{random.choice('!@#$%')}"

def load_names() -> list:
    """Pehli baar names.txt banao agar nahi hai toh"""
    if not os.path.exists("names.txt"):
        with open("names.txt", "w") as f:
            f.write("First Last\n")
            for fn in FIRST_NAMES:
                for ln in random.sample(LAST_NAMES, 3):
                    f.write(f"{fn} {ln}\n")
    with open("names.txt") as f:
        lines = [l.strip().split() for l in f if l.strip()]
    return [(l[0], l[1]) for l in lines[1:] if len(l) >= 2]

# The step "Next"/"Create" button. Outlook's current signup uses
# <button type=submit data-testid=primaryButton>; older flows used
# <input type=submit>. Match either so each step advances.
PRIMARY_BTN = ('button[data-testid="primaryButton"], '
               'button[type="submit"], input[type="submit"]')

MONTHS = ["January", "February", "March", "April", "May", "June",
          "July", "August", "September", "October", "November", "December"]

def select_dob_dropdown(page, dropdown_id, option_text, timeout=10000):
    """Pick a value from a Fluent combobox (DOB month/day). A floating <label>
    overlaps the button so a normal click is intercepted — force it open, then
    click the option by exact text."""
    page.click(f'#{dropdown_id}', force=True)
    page.wait_for_selector('[role="option"]', timeout=timeout)
    time.sleep(0.4)
    for o in page.query_selector_all('[role="option"]'):
        if (o.inner_text() or '').strip() == option_text:
            try:
                o.click()
            except Exception:
                o.click(force=True)
            return
    raise Exception(f"DOB option {option_text!r} not found in #{dropdown_id}")

CAPTCHA_SELS = [
    'iframe[src*="hsprotect.net"]',
    '[data-testid="humanCaptchaIframe"]',
    'iframe[src*="funcaptcha"]',
    'iframe[src*="arkoselabs"]',
]
BLOCK_PHRASES = [
    "account creation has been blocked",
    "we've blocked",
    "account has been blocked",
    "creation is blocked",
    "too many accounts",
    "suspicious activity",
]

def captcha_present(page) -> bool:
    for sel in CAPTCHA_SELS:
        try:
            if page.query_selector(sel):
                return True
        except Exception:
            pass
    # Fallback: "prove you're human" heading
    try:
        el = page.query_selector('h1[data-testid="title"]')
        if el and "prove you" in (el.inner_text() or "").lower():
            return True
    except Exception:
        pass
    return False

def inbox_ready(page) -> bool:
    """Check if Junk Email folder visible (we're in mailbox)."""
    for sel in [
        'div[data-folder-name="junk email"]',
        '[title*="Junk Email"]',
        'span:has-text("Junk Email")',
    ]:
        try:
            if page.query_selector(sel):
                return True
        except Exception:
            pass
    return False

# ── Outlook signup flow (returns True on success, False if email taken) ─
def outlook_signup(page, email, password, first, last, month, day, year) -> bool:
    """Perform all steps. Raises Exception on block/timeout."""
    # Step 1: Email
    print("  [OL1] Filling email...")
    page.wait_for_selector('input[type="email"]', timeout=30000)
    page.fill('input[type="email"]', email.split("@")[0])
    page.click(PRIMARY_BTN)
    time.sleep(2)

    body = page.inner_text("body").lower()
    if any(p in body for p in ["already taken", "not available", "someone already"]):
        return False

    # Step 2: Password
    print("  [OL2] Filling password...")
    page.wait_for_selector('input[type="password"]', timeout=30000)
    page.fill('input[type="password"]', password)
    page.click(PRIMARY_BTN)
    time.sleep(2)

    # Step 3: DOB — month/day are Fluent dropdowns, year is a number input
    print("  [OL3] Filling DOB...")
    page.wait_for_selector('#BirthMonthDropdown', timeout=20000)
    select_dob_dropdown(page, 'BirthMonthDropdown', MONTHS[month - 1])
    time.sleep(0.6)
    select_dob_dropdown(page, 'BirthDayDropdown', str(day))
    time.sleep(0.6)
    page.fill('input[name="BirthYear"]', str(year))
    time.sleep(0.6)
    page.click(PRIMARY_BTN)
    time.sleep(2)

    # Step 4: Name
    print("  [OL4] Filling name...")
    try:
        page.wait_for_selector('input#firstNameInput', timeout=20000)
        page.fill('input#firstNameInput', first)
        page.fill('input#lastNameInput', last)
        page.click(PRIMARY_BTN)
    except Exception:
        page.fill('input[aria-label="First name"]', first)
        page.fill('input[aria-label="Last name"]', last)
        page.click(PRIMARY_BTN)

    # Step 5: Captcha / block check + manual solve
    print("  [OL5] Captcha/block check...")
    MANUAL_TIMEOUT = 300  # 5 minutes

    for _ in range(6):
        captcha_found = False
        for _ in range(30):
            body_lower = page.inner_text("body").lower()
            if any(p in body_lower for p in BLOCK_PHRASES):
                raise Exception("AccountBlocked")
            if captcha_present(page):
                captcha_found = True
                break
            if inbox_ready(page):
                print("  [i] Inbox ready!")
                return True
            time.sleep(3)

        if not captcha_found:
            break  # no captcha, proceed

        # Manual captcha solving
        print("  [OL5] *** CAPTCHA DETECTED — solve manually on the phone ***")
        print(f"  [OL5] Script will wait up to {MANUAL_TIMEOUT}s...")
        deadline = time.time() + MANUAL_TIMEOUT
        solved = False
        while time.time() < deadline:
            remaining = int(deadline - time.time())
            print(f"  [OL5] {remaining}s remaining   ", end="\r", flush=True)
            if not captcha_present(page):
                print("\n  [OL5] Captcha solved!")
                solved = True
                break
            if inbox_ready(page):
                print("\n  [i] Inbox ready after captcha!")
                return True
            time.sleep(2)
        if not solved:
            raise Exception("CaptchaDetected")

        # quick check after solve
        for _ in range(10):
            if inbox_ready(page):
                print("  [i] Inbox ready after captcha!")
                return True
            time.sleep(1)

    # Step 6: Wait for inbox
    print("  [OL6] Waiting for inbox (max 3 min)...")
    for _ in range(60):
        if inbox_ready(page):
            return True
        time.sleep(3)

    print("  [!] Inbox timeout — assuming success")
    return True

# ── OTP from Outlook Junk folder ────────────────────────────────────────
def get_otp_from_outlook(context, email: str) -> str | None:
    print("  [OTP] Reading OTP from Junk folder...")
    otp_page = context.new_page()
    try:
        otp_page.goto("https://outlook.live.com/mail/0/inbox",
                      wait_until="domcontentloaded", timeout=30000)
        time.sleep(3)

        # Click Junk Email
        junk_clicked = False
        for sel in [
            'div[data-folder-name="junk email"]',
            '[title*="Junk Email"]',
            'span:has-text("Junk Email")',
        ]:
            el = otp_page.query_selector(sel)
            if el:
                el.click()
                junk_clicked = True
                time.sleep(2)
                break

        for attempt in range(40):
            if attempt > 0 and attempt % 5 == 0:
                try:
                    otp_page.reload(wait_until="domcontentloaded")
                    time.sleep(2)
                    if junk_clicked:
                        for sel in ['div[data-folder-name="junk email"]']:
                            el = otp_page.query_selector(sel)
                            if el:
                                el.click()
                                time.sleep(2)
                                break
                except Exception:
                    pass

            # Subject lines
            for subj_el in otp_page.query_selector_all('span.TtcXM'):
                try:
                    txt = subj_el.inner_text().strip()
                    m = re.search(r'^([A-Za-z0-9]{6})\s+is your Bright Data access code', txt)
                    if m:
                        print(f"  [+] OTP found: {m.group(1)}")
                        return m.group(1)
                except Exception:
                    continue

            # Body of emails from BrightData
            for sender in otp_page.query_selector_all('span[title="noreply@bright-notice.com"]'):
                try:
                    sender.click()
                    time.sleep(2)
                    body = otp_page.inner_text("body")
                    m = re.search(r'([A-Za-z0-9]{6})\s+is your Bright Data access code', body)
                    if m:
                        print(f"  [+] OTP found: {m.group(1)}")
                        return m.group(1)
                except Exception:
                    pass

            print(f"  [.] Waiting for OTP ({attempt+1}/40)...")
            time.sleep(3)
    except Exception as e:
        print(f"  [!] OTP error: {e}")
    finally:
        try:
            otp_page.close()
        except Exception:
            pass
    return None

# ── BrightData signup (inline — drives brightdata.com on the phone via CDP) ─
def _bd_captcha_showing(page) -> bool:
    """Cloudflare Turnstile renders in a child frame whose <iframe src> attribute
    isn't reliably queryable, so detect it via the frame list instead."""
    for fr in page.frames:
        u = fr.url or ""
        if "challenges.cloudflare.com" in u or "turnstile" in u:
            return True
    return False

def _bd_token_present(page) -> bool:
    """True once Cloudflare Turnstile has produced its response token (i.e.
    solved). This is the reliable 'done' signal — the frame can linger."""
    el = page.query_selector('input[name="cf-turnstile-response"]')
    return bool(el and (el.get_attribute("value") or "").strip())

def bd_wait_captcha(page, label, timeout=300) -> bool:
    """Cloudflare Turnstile must be solved by hand on the phone. Wait until a
    token appears. No-op if no captcha frame is showing."""
    if not _bd_captcha_showing(page):
        return True
    print(f"  [BD] *** Cloudflare captcha ({label}) — solve it on the phone ***")
    deadline = time.time() + timeout
    while time.time() < deadline:
        if _bd_token_present(page):
            print("\n  [BD] captcha solved (token present)")
            return True
        print(f"  [BD] {int(deadline - time.time())}s left   ", end="\r", flush=True)
        time.sleep(2)
    print("\n  [BD] captcha timeout")
    return False

def _bd_fill_first(page, selectors, value, label, timeout=12000) -> bool:
    for sel in selectors:
        try:
            page.wait_for_selector(sel, timeout=timeout)
            page.fill(sel, value)
            print(f"  [BD] {label} filled ({sel})")
            return True
        except Exception:
            continue
    print(f"  [BD] !! {label}: no selector matched")
    return False

def _bd_click_first(page, selectors, label, timeout=8000) -> bool:
    for sel in selectors:
        try:
            page.wait_for_selector(sel, timeout=timeout)
            page.click(sel)
            print(f"  [BD] {label} clicked ({sel})")
            return True
        except Exception:
            continue
    print(f"  [BD] !! {label}: no button matched")
    return False

def bd_fill_otp(page, code) -> bool:
    """Enter the 6-char code into BrightData's OTP boxes (per-digit) or a single
    number input as a fallback."""
    boxes = page.query_selector_all('[data-testid^="otp_input-"] input')
    if len(boxes) >= len(code):
        for i, ch in enumerate(code):
            boxes[i].fill(ch)
        return True
    el = (page.query_selector('[data-testid="otp_input-0"] input')
          or page.query_selector('input[type="number"]'))
    if el:
        el.click()
        page.keyboard.type(code, delay=120)
        return True
    return False

def brightdata_signup_inline(ctx, email, password) -> bool:
    """Create a BrightData account on the phone's Chrome, reading the email
    verification code from the just-created Outlook inbox (same context).
    Cloudflare captchas are solved manually on the phone. True on success."""
    page = ctx.new_page()
    print("  [BD1] opening brightdata signup...")
    page.goto("https://brightdata.com/?hs_signup=1",
              wait_until="domcontentloaded", timeout=60000)
    time.sleep(4)

    # The signup popup (?hs_signup=1) renders inconsistently (VWO A/B test +
    # Cloudflare) — sometimes only a bare email field shows. Reload until the
    # Create Account button actually appears, then fill + solve captcha + click.
    for attempt in range(3):
        _bd_fill_first(page, ['input[name="email"]', 'input[type="email"]'],
                       email, "email", timeout=20000)
        time.sleep(1)
        try:
            page.wait_for_selector('button:has-text("Create Account")', timeout=15000)
            break
        except Exception:
            print(f"  [BD] signup form not ready (try {attempt + 1}/3), reloading...")
            page.goto("https://brightdata.com/?hs_signup=1",
                      wait_until="domcontentloaded", timeout=60000)
            time.sleep(4)
    bd_wait_captcha(page, "signup form")
    if not _bd_click_first(page, ['button:has-text("Create Account")',
                                  'button.hs-button.primary.large',
                                  'input[type="submit"][value="Create Account"]'],
                           "Create Account", timeout=30000):
        return False

    # Password page (/cp/signup) — poll up to ~2 min
    print("  [BD2] waiting for password field...")
    got_pwd = False
    for _ in range(24):
        if page.query_selector('input#password'):
            got_pwd = True
            break
        time.sleep(5)
    if not got_pwd:
        print("  [BD] !! password field never appeared")
        return False
    _bd_fill_first(page, ['input#password', 'input[placeholder="Set your password"]'],
                   password, "password")
    _bd_fill_first(page, ['input#password_confirm',
                          'input[placeholder="Confirm your password"]'],
                   password, "confirm password")
    bd_wait_captcha(page, "password form")
    _bd_click_first(page, ['button[type="submit"].signup',
                           'button[type="submit"]:has-text("Sign up")',
                           'button:has-text("Sign up")'], "Sign up")

    # OTP box → read code from Outlook → enter it
    print("  [BD3] waiting for OTP box...")
    otp_ready = False
    for _ in range(20):  # ~60s
        if (page.query_selector('[data-testid="otp_input-0"] input')
                or page.query_selector('input[type="number"]')):
            otp_ready = True
            break
        time.sleep(3)
    if not otp_ready:
        print("  [BD] !! OTP box never appeared")
        return False
    code = get_otp_from_outlook(ctx, email)
    if not code:
        print("  [BD] !! BrightData OTP not found in Outlook")
        return False
    if not bd_fill_otp(page, code):
        print("  [BD] !! could not enter OTP")
        return False
    print(f"  [BD] OTP {code} entered")
    time.sleep(2)
    _bd_click_first(page, ['button[type="submit"]:has-text("Verify")',
                           'button[type="submit"]:has-text("Continue")',
                           'button[type="submit"]'], "verify OTP", timeout=5000)

    # Landed in the control panel?
    for _ in range(20):
        if "/cp" in page.url:
            print(f"  [BD] reached dashboard: {page.url[:60]}")
            return True
        time.sleep(3)
    print(f"  [BD] finished (final url: {page.url[:60]})")
    return True

# ── CDP helpers ─────────────────────────────────────────────────────────
def adb_base(serial):
    cmd = ["adb"]
    if serial:
        cmd += ["-s", serial]
    return cmd

def ensure_forward(serial, port, socket):
    base = adb_base(serial)
    try:
        out = subprocess.check_output(
            base + ["shell", "cat", "/proc/net/unix"],
            stderr=subprocess.STDOUT, text=True, timeout=10,
        )
        if socket not in out:
            print(f"!! Socket '@{socket}' not found on device.")
            print("   Open Chrome on the phone at least once, then retry.")
            sys.exit(1)
    except subprocess.CalledProcessError as e:
        print("!! adb shell failed:\n" + e.output)
        sys.exit(1)
    except FileNotFoundError:
        print("!! `adb` not found on PATH.")
        sys.exit(1)

    subprocess.run(
        base + ["forward", f"tcp:{port}", f"localabstract:{socket}"],
        check=True, stdout=subprocess.DEVNULL,
    )

def cdp_get(port, path):
    url = f"http://localhost:{port}{path}"
    with urllib.request.urlopen(url, timeout=10) as r:
        return r.read().decode()

# ── "Incognito" for Android Chrome ─────────────────────────────────────
# Real incognito is NOT reachable here: Android Chrome doesn't support the CDP
# Target.createBrowserContext call (Playwright's browser.new_context() throws
# "Failed to create browser context"), and it blocks external intents from
# opening an incognito tab. So --incognito instead wipes the cookies for the
# Microsoft/Outlook domains below, giving each signup a fresh logged-out start
# without touching any other site's login on the phone.
MS_DOMAINS = [
    "live.com", "microsoft.com", "microsoftonline.com",
    "outlook.com", "office.com", "bing.com", "msn.com",
]

def fresh_ms_session(ctx):
    """Clear cookies for Microsoft/Outlook domains only. Returns (before, after)."""
    cookies = ctx.cookies()
    before = len(cookies)
    targets = set()
    for c in cookies:
        dom = c.get("domain", "").lstrip(".")
        if any(dom == d or dom.endswith("." + d) for d in MS_DOMAINS):
            targets.add(c["domain"])  # exact domain string for the filter
    for d in targets:
        try:
            ctx.clear_cookies(domain=d)
        except Exception:
            pass
    after = len(ctx.cookies())
    print(f"  [incognito] Microsoft session cleared: {before} -> {after} cookies")
    return before, after

# ── Original commands (unchanged except incognito handling) ────────────
def pick_active_page(ctx):
    pages = ctx.pages
    if not pages:
        return ctx.new_page()
    for page in reversed(pages):
        try:
            if page.evaluate("document.visibilityState") == "visible":
                return page
        except Exception:
            continue
    return pages[-1]

def connect(args):
    """Return (browser, context, page) connected to phone Chrome."""
    ensure_forward(args.serial, args.port, args.socket)
    browser = cdp.connect(args.port)
    # Android Chrome only ever exposes the single default context.
    ctx = browser.contexts[0]
    if args.incognito:
        fresh_ms_session(ctx)
    page = pick_active_page(ctx)
    return browser, ctx, page

def cmd_forward(args):
    ensure_forward(args.serial, args.port, args.socket)
    print(f">> Forwarded tcp:{args.port} -> @{args.socket}")

def cmd_info(args):
    ensure_forward(args.serial, args.port, args.socket)
    print(cdp_get(args.port, "/json/version"))

def cmd_tabs(args):
    ensure_forward(args.serial, args.port, args.socket)
    data = json.loads(cdp_get(args.port, "/json"))
    pages = [t for t in data if t.get("type") == "page"]
    print(f"{len(pages)} tab(s):")
    for t in pages:
        print(f"  - {t.get('title','')[:50]!r}  {t.get('url','')[:70]}")

def cmd_open(args):
    browser, ctx, _ = connect(args)
    page = ctx.new_page()
    page.goto(args.url, wait_until="domcontentloaded")
    print(f">> Opened: {page.url}\n   title: {page.title()}")
    browser.close()

def cmd_goto(args):
    browser, ctx, page = connect(args)
    page.goto(args.url, wait_until="domcontentloaded")
    print(f">> {page.url}\n   title: {page.title()}")
    browser.close()

def cmd_eval(args):
    browser, ctx, page = connect(args)
    result = page.evaluate(args.js)
    print(result)
    browser.close()

def cmd_text(args):
    browser, ctx, page = connect(args)
    txt = page.evaluate("document.body ? document.body.innerText : ''")
    print(txt[:2000])
    browser.close()

def cmd_shot(args):
    ensure_forward(args.serial, args.port, args.socket)
    subprocess.run(
        adb_base(args.serial) + ["shell", "monkey", "-p", "com.android.chrome",
                                 "-c", "android.intent.category.LAUNCHER", "1"],
        stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
    )
    time.sleep(1.2)
    with open(args.file, "wb") as f:
        subprocess.run(
            adb_base(args.serial) + ["exec-out", "screencap", "-p"],
            check=True, stdout=f,
        )
    print(f">> Saved {args.file}")

def cmd_screen(args):
    ensure_forward(args.serial, args.port, args.socket)
    with open(args.file, "wb") as f:
        subprocess.run(
            adb_base(args.serial) + ["exec-out", "screencap", "-p"],
            check=True, stdout=f,
        )
    print(f">> Saved {args.file}")

def cmd_fill(args):
    browser, ctx, page = connect(args)
    page.fill(args.selector, args.value)
    print(f">> Filled {args.selector!r}")
    browser.close()

def cmd_click(args):
    browser, ctx, page = connect(args)
    page.click(args.selector)
    print(f">> Clicked {args.selector!r}")
    browser.close()

# ── Signup command ──────────────────────────────────────────────────────
def cmd_signup(args):
    """Outlook account create karega (incognito flag ke saath)."""
    ensure_forward(args.serial, args.port, args.socket)
    names = load_names()
    first, last = random.choice(names)
    email = gen_email(first, last)
    password = gen_password(first)
    month = random.randint(1, 12)
    day = random.randint(1, 28)
    year = random.randint(1985, 2005)
    print(f"\n🎯 Creating: {first} {last} | {email}")

    browser = cdp.connect(args.port)
    ctx = browser.contexts[0]
    # "Incognito" = clear any existing Microsoft/Outlook login before signup.
    if args.incognito:
        fresh_ms_session(ctx)
    page = ctx.new_page()

    try:
        page.goto("https://outlook.live.com/mail/?prompt=create_account",
                  wait_until="domcontentloaded", timeout=60000)

        # Try signup (email taken pe retry)
        success = False
        for _ in range(5):
            try:
                result = outlook_signup(page, email, password, first, last, month, day, year)
                if result is False:  # email taken
                    email = gen_email(first, last)
                    print(f"  Email taken, new: {email}")
                    page.goto("https://outlook.live.com/mail/?prompt=create_account",
                              wait_until="domcontentloaded")
                    continue
                success = True
                break
            except Exception as e:
                err = str(e)
                if "CaptchaDetected" in err:
                    print("  [!] Captcha timeout — skipping")
                    return
                elif "AccountBlocked" in err:
                    print("  [!] Account creation blocked by Microsoft")
                    return
                else:
                    print(f"  [!] Signup error: {e}")
                    return

        if not success:
            return

        # Save Outlook creds
        os.makedirs("outlook_cookies", exist_ok=True)
        with open("outlook_accounts.txt", "a") as f:
            f.write(f"{email}:{password}\n")
        cookie_file = os.path.join("outlook_cookies", f"cookies_{email}.json")
        with open(cookie_file, "w") as f:
            json.dump(ctx.cookies(), f)
        print(f"  ✅ Outlook saved: {email}")

        # Optional BrightData signup (inline, on the phone; OTP read at the
        # right moment from this same Outlook inbox)
        if args.brightdata:
            print("  [BD] Starting BrightData signup (inline)...")
            if brightdata_signup_inline(ctx, email, password):
                with open("brightdata_accounts.txt", "a") as f:
                    f.write(f"{email}:{password}\n")
                print(f"  ✅ BrightData saved: {email}")
            else:
                print("  [!] BrightData signup did not complete")
        else:
            print("  [i] BrightData skipped (use --brightdata to enable)")

    finally:
        browser.close()

# ── CLI ──────────────────────────────────────────────────────────────────

_SUBCOMMANDS = frozenset([
    'forward', 'info', 'tabs', 'open', 'goto', 'eval',
    'text', 'shot', 'screen', 'fill', 'click', 'signup',
])
_HOIST_FLAGS = {'--incognito'}

def _normalize_argv(argv):
    """Allow --incognito anywhere (before or after subcommand) by hoisting it
    before the subcommand so the parent parser always sees it."""
    pre, sub, hoisted = [], [], []
    past_sub = False
    for a in argv:
        if not past_sub and a in _SUBCOMMANDS:
            past_sub = True
            sub.append(a)
        elif past_sub and a in _HOIST_FLAGS:
            hoisted.append(a)
        else:
            (sub if past_sub else pre).append(a)
    return pre + hoisted + sub


def main():
    p = argparse.ArgumentParser(description="Drive phone Chrome over CDP.")
    p.add_argument("--serial", default=None, help="adb serial, e.g. 192.168.1.2:5555")
    p.add_argument("--port", type=int, default=9222)
    p.add_argument("--socket", default="chrome_devtools_remote")
    p.add_argument("--incognito", action="store_true", help="Use incognito/private mode")
    sub = p.add_subparsers(dest="cmd", required=True)

    sub.add_parser("forward").set_defaults(func=cmd_forward)
    sub.add_parser("info").set_defaults(func=cmd_info)
    sub.add_parser("tabs").set_defaults(func=cmd_tabs)

    sp = sub.add_parser("open"); sp.add_argument("url"); sp.set_defaults(func=cmd_open)
    sp = sub.add_parser("goto"); sp.add_argument("url"); sp.set_defaults(func=cmd_goto)
    sp = sub.add_parser("eval"); sp.add_argument("js"); sp.set_defaults(func=cmd_eval)
    sub.add_parser("text").set_defaults(func=cmd_text)
    sp = sub.add_parser("shot"); sp.add_argument("file"); sp.set_defaults(func=cmd_shot)
    sp = sub.add_parser("screen"); sp.add_argument("file"); sp.set_defaults(func=cmd_screen)
    sp = sub.add_parser("fill"); sp.add_argument("selector"); sp.add_argument("value"); sp.set_defaults(func=cmd_fill)
    sp = sub.add_parser("click"); sp.add_argument("selector"); sp.set_defaults(func=cmd_click)

    sp_signup = sub.add_parser("signup", help="Create Outlook account (and optionally BrightData)")
    sp_signup.add_argument("--brightdata", action="store_true", help="Also perform BrightData signup")
    sp_signup.set_defaults(func=cmd_signup)

    args = p.parse_args(_normalize_argv(sys.argv[1:]))
    args.func(args)

if __name__ == "__main__":
    main()
