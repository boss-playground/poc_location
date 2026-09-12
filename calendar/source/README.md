# My calendar — Outlook + browser events

A standalone static month-calendar page at `/calendar/`. Local events persist in the same browser; Outlook events are read and written through Microsoft Graph. No backend or client secret is used.

## Plum & Gold interface

The calendar uses a generated abstract violet ribbon background, translucent plum panels, modern sans-serif headings, gold actions, and distinct lilac/local and gold/Outlook events. The compact sidebar has no duplicate mini calendar; the calendar fills the desktop workspace with a 24px bottom gutter. The header's pause button stops ambient motion; Reduce Motion disables animation automatically. Local data, Microsoft configuration and event operations are unchanged. The purple/gold direction is SCB-inspired, without a bank logo or affiliation.

The production asset is `public/assets/plum-gold.png`; the built-in Image Gen concept and implementation decisions are in `design/plum-gold-concept.png` and `design/plum-gold.md`. No extra graphics runtime or third-party font request is needed.

### Popup cancellation

MSAL 5.21.0 does not observe a manually closed popup while waiting for its redirect bridge. `src/outlook-popup-client.mjs` observes closure and cancels only the request-specific bridge, allowing MSAL to release its own interaction lock. It does not clear account/token caches or local events. This compatibility hook must be revalidated when upgrading MSAL.

Before the bridge starts, a blank popup can be waiting on authority discovery. Discovery GETs have a 10-second abort deadline per request (two sequential lookups can take about 20 seconds); early closure waits for this bounded cleanup, rather than immediately unlocking an unfinished request. Token POSTs retain MSAL's normal network implementation.

If an old or unknown interaction lock survives a reload, it is intentionally not removed: another request might own it. Open the calendar URL in a fresh tab (not Duplicate tab) and reconnect. The new close/reconnect behavior is covered by real-MSAL browser tests with mocked identity endpoints; live Microsoft login/consent remains a separate manual check.

## Run

Use Node.js 22.12+ (verified here with Node 25.9.0).

```sh
cd /Users/bossbew/Documents/project/poc_location/calendar
npm install
npm run dev
```

Open **http://localhost:5173/calendar/**. The server uses a strict port so the Microsoft redirect URI stays predictable. Use this exact host; `localhost` and `127.0.0.1` have separate browser storage and different OAuth redirect configuration.

```sh
npm test
npm run build
npm run preview
```

The production output is `calendar/dist/`. Serve its contents beneath `/calendar/` on an HTTPS static host. `index.html`, `redirect.html` and `logout.html` must be served as actual files. `file://` is unsupported.

## ทดลอง localStorage ได้ทันที

1. กด **New event** หรือคลิกวันที่ในปฏิทิน
2. กรอกชื่อ วัน และเวลา เลือก **This calendar** แล้วกด **Save event**
3. Reload หน้าเพื่อดูว่า event ยังอยู่
4. คลิก event เพื่อแก้ไข หรือเลือก **Delete event → Confirm delete**

ข้อมูลเก็บใน browser นี้บน origin เดิม ภายใต้ key `outlook-calendar-poc.local-events.v1` เท่านั้น ไม่ได้ส่งไป Outlook และไม่ข้ามเครื่อง การล้างข้อมูลเว็บไซต์/ใช้ private browsing อาจทำให้ข้อมูลหาย จึงไม่ควรใช้เป็นที่เก็บข้อมูลสำคัญเพียงแห่งเดียว

## ตั้งค่า Outlook ส่วนตัว

คุณต้องมี Outlook.com/Hotmail mailbox และสิทธิ์สร้าง app registration ใน Microsoft Entra tenant การมีบัญชี Outlook ส่วนตัวเพียงอย่างเดียวอาจยังไม่ให้สิทธิ์เข้าเมนู App registrations หากยังไม่มี tenant ให้ทำตาม [Microsoft registration prerequisites](https://learn.microsoft.com/en-us/graph/auth-register-app-v2) หรือใช้ tenant ที่คุณมีสิทธิ์อยู่แล้ว

1. เปิด [Microsoft Entra admin center](https://entra.microsoft.com/)
2. ไปที่ **App registrations → New registration** ตั้งชื่อ `Outlook Calendar POC`
3. เลือก supported account types เป็น **Personal Microsoft accounts only** (ชื่อใน portal อาจแสดง Personal Microsoft accounts)
4. ใน **Authentication → Add a platform → Single-page application** เพิ่ม redirect URI ทั้งสองค่า (สำหรับ login และ logout):

   ```text
   http://localhost:5173/calendar/redirect.html
   http://localhost:5173/calendar/logout.html
   ```

5. ใน **API permissions → Add a permission → Microsoft Graph → Delegated permissions** เพิ่ม `User.Read` และ `Calendars.ReadWrite`
6. ไม่ต้องสร้าง client secret และไม่ต้องเปิด implicit grant
7. คัดลอก **Application (client) ID** จาก Overview
8. กลับมาหน้าปฏิทิน กด **Connect Outlook** วาง Client ID แล้วกด **Use this application** หน้า reload เพื่อเตรียม MSAL
9. กด **Connect Outlook** อีกครั้งเพื่อเปิดหน้าล็อกอิน Microsoft และยินยอมให้เข้าถึง Calendar

Client ID เป็นค่าที่เปิดเผยได้ ไม่ใช่รหัสผ่าน ช่องตั้งค่าเก็บ ID ไว้ใน `sessionStorage` ของ tab นี้ ถ้าต้องการตั้งถาวร ให้ใส่ค่า `clientId` ใน `src/config.mjs` แล้ว build ใหม่ อย่าใส่รหัสผ่านหรือ client secret

เมื่อ deploy ต้องเพิ่ม `https://YOUR-HOST/calendar/redirect.html` และ `https://YOUR-HOST/calendar/logout.html` เป็น SPA redirect URI ด้วย ทั้งสองหน้าใช้ MSAL v5 redirect bridge จึงห้ามแทนด้วยหน้าเปล่าหรือ redirect กลับ home อัตโนมัติ ดู [Microsoft redirect bridge documentation](https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/redirect-bridge.md)

การอ่าน/เขียน calendar ปกติผ่าน delegated Graph API ไม่ต้องซื้อ API plan เพิ่มสำหรับ POC นี้ ส่วน Entra tenant onboarding อาจต้องมี Azure billing account/บัตรยืนยันตัวตน และบริการ Azure อื่นมีเงื่อนไขค่าใช้จ่ายแยก ดู [Entra ID Free](https://learn.microsoft.com/en-us/azure/cost-management-billing/manage/microsoft-entra-id-free)

## เปรียบเทียบสองแหล่ง

| Destination | เก็บที่ไหน | ผลของการแก้ไข/ลบ |
|---|---|---|
| This calendar — สีม่วง | localStorage ของ browser | มีผลเฉพาะ browser นี้ |
| Outlook — สีทอง | Default calendar ของบัญชี Microsoft | ส่งคำสั่งไป Outlook จริง |

เลือก destination ตอนสร้าง event หลังสร้าง source จะคงเดิม การปิด checkbox แค่ซ่อนรายการ ไม่ลบข้อมูล Event ที่เพิ่มภายในจะไม่ถูก upload ทั้งหมดหลังเชื่อมบัญชี

Sync จะอ่าน Outlook ตอน login, reload, เปลี่ยนเดือน, หลังเขียน Outlook สำเร็จ หรือกด **Sync Outlook** ไม่ sync เบื้องหลัง เมื่อแก้จาก Outlook บนอีกหน้าจอ ให้กด Sync Outlook เพื่อดึงผลล่าสุด

Outlook token จัดการโดย MSAL ใน `sessionStorage` ไม่มีการนำ token ไปเก็บใน local event records การ disconnect เก็บ local events ไว้

## Scope and time behavior

- Month view and default Outlook calendar only; no shared calendars, attendee editor or recurrence editor.
- All displayed timed events use **Bangkok / UTC+07:00** regardless of the computer timezone.
- Time inputs support seconds; changing only a title preserves the event's original seconds.
- All-day editor end date is inclusive for people (12–12 September means one day); storage and Graph use exclusive end (13 September).
- Existing all-day Outlook items retain their source calendar dates/timezone; their dates do not shift when projected to Bangkok.
- Recurring Outlook occurrences/series are shown read-only. Manage them in Outlook.
- Editing or deleting an existing Outlook meeting may notify attendees or send cancellation messages. The dialog identifies these items. Use a personal test appointment without attendees for initial tests.
- Offline local CRUD continues. Outlook errors keep the editor draft available. No automated write retries. If a request loses its connection while saving, check Outlook before retrying because the server may already have applied it.
- This POC has no cross-device local-event storage or conflict-merge engine. Refresh before editing an Outlook appointment changed elsewhere.

## Validation and remaining live check

`npm test` exercises domain dates/validation, storage CRUD and failure preservation, source dispatch/account races, MSAL adapter behavior, Graph payloads/pagination/errors/timezone handling.

`npm run test:e2e` uses Playwright with the installed Google Chrome in a fresh isolated profile. It checks actual DOM interactions, local persistence and source filters, all-day entry, mobile overflow, setup, and Graph requests through a mocked Microsoft auth/network boundary. Install Google Chrome first or change the Playwright channel to your installed browser. Screenshots contain synthetic test data and are saved under `/tmp/calendar-poc-*.png`.

**Live Outlook sign-in and real Graph writes are not verified until you supply your Client ID and sign in.** Mock tests do not prove tenant registration, Microsoft consent, popup policy, or live mailbox behavior.

Live acceptance: create an appointment named `POC Outlook test` without attendees, choose Outlook, verify it at [Outlook Calendar](https://outlook.live.com/calendar/), change its time here and verify again, delete that test appointment here and verify its removal. Also create a local-only appointment and confirm it never appears in Outlook.

## Troubleshooting

- **Redirect mismatch / AADSTS50011:** compare the exact redirect URI shown in connection settings with the SPA registration. Include `/calendar/redirect.html` and the port.
- **Account type not supported:** register personal Microsoft accounts and use Outlook.com/Hotmail with the configured `consumers` authority.
- **Popup blocked or cancelled:** allow popups for localhost, then reconnect. Cancellation does not affect local events.
- **Permission denied / 403:** verify delegated `Calendars.ReadWrite`, reconnect, and grant consent.
- **Session expired:** disconnect and connect again; background reads intentionally never open unexpected login popups.
- **Local storage unreadable:** saved data is preserved. Check browser storage/privacy settings and inspect/export the existing namespaced storage value before repairing it.

## Primary references

- [MSAL Browser initialization and popup redirect bridge](https://learn.microsoft.com/en-us/entra/msal/javascript/browser/initialization)
- [Graph calendarView](https://learn.microsoft.com/en-us/graph/api/calendar-list-calendarview?view=graph-rest-1.0)
- [Create event](https://learn.microsoft.com/en-us/graph/api/user-post-events?view=graph-rest-1.0)
- [Update event](https://learn.microsoft.com/en-us/graph/api/event-update?view=graph-rest-1.0)
- [Delete event and cancellation behavior](https://learn.microsoft.com/en-us/graph/api/event-delete?view=graph-rest-1.0)
