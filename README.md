# Autofill Google Form

A Google Apps Script web app that looks up the signed-in student in a
spreadsheet and forwards them to a Google Form already filled in with their
own details.

One deployment serves **any number of forms** — each one is a row in the
`Forms` tab, selected with `?form=<key>`.

---

## Setup

1. Create a Google Sheet → **Extensions → Apps Script**.
2. Paste `Code.gs` into the editor and save.
3. **Deploy → New deployment → Web app**, and set:
   - **Execute as:** `User accessing the web app` ← *required*, otherwise the
     script cannot see who the visitor is and everybody gets a blank form.
   - **Who has access:** everyone who should be able to use the link.

### The `SYNCDATA` tab (the roster)

One row per student. Headers are matched by name, so the order does not
matter; when a header is not recognised the script falls back to these fixed
positions:

| Column | Field           |
| ------ | --------------- |
| D      | title           |
| E      | first_name      |
| F      | last_name       |
| G      | nickname        |
| K      | docchula_email  |
| L      | phone_number    |
| M      | line_id         |
| N      | line_display    |

Recognised header names include `first_name`, `last_name`, `nickname`,
`phone_number`, `line_id`, `line_display`, `title`, `year`, and Thai
equivalents (`ชื่อ`, `นามสกุล`, `ชื่อเล่น`, `เบอร์โทร`, `คำนำหน้า`, `ชั้นปี`).

A student is matched on **any** email-like column, so extra address columns
are safe to add.

### The `Forms` tab (the form registry)

| A (key) | B (template_url)        | C (blank_url)              |
| ------- | ----------------------- | -------------------------- |
| Test1   | pre-filled link, form 1 | plain link, form 1          |
| Test2   | pre-filled link, form 2 | plain link, form 2          |

- **key** — what goes in `?form=`. Case and spaces are ignored.
- **template_url** — from the form's **⋮ → Get pre-filled link**. It must
  contain `?entry.…`; the plain form link will not work.
- **blank_url** — where visitors who are not on the roster are sent.

Then hand out one link per form:

```
https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?form=Test1&openExternalBrowser=1
https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?form=Test2&openExternalBrowser=1
```

`openExternalBrowser=1` makes LINE open the link in the real browser, where
the student is signed in to Google.

---

## Telling the script which answer goes where

Use either method, or mix them.

### 1. Placeholder text (no extra setup)

When building the pre-filled link, type a marker word into each question:

| Type this          | Gets replaced with          |
| ------------------ | --------------------------- |
| `DummyName`        | full name (first + last)    |
| `DummyName` + `DummyLastName` | first name / last name, when the form asks separately |
| `DummyFullName`    | full name                   |
| `DummyTitle`       | title (คำนำหน้า)            |
| `DummyNickname`    | nickname                    |
| `0XX-XXX-XXXX`     | phone, as `081-234-5678`    |
| `DummyLineID`      | Line ID                     |
| `DummyLineDisplay` | Line display name           |
| `DummyEmail`       | email                       |
| `DummyYear`        | year (ชั้นปี)               |
| `6XXXXXXX30`       | *nothing* — student ID is deliberately left blank |

Each form may use its own subset. A question whose marker cannot be filled
in (no data on file, or a marker the script does not know) is **left empty**
rather than showing `DummyNickname` to the student.

### 2. Entry-id mapping

Dropdowns, multiple choice and linear scale questions cannot hold marker
text. For those, add columns from **D onwards** to the `Forms` tab: the
**header** is the field name and the **cell** is that question's entry id.

| A     | B            | C         | D (fullname)       | E (nickname)       |
| ----- | ------------ | --------- | ------------------ | ------------------ |
| Test2 | *(link)*     | *(link)*  | entry.1234567890   | entry.9876543210   |

A single column headed `mapping` also works:
`fullname=entry.1234567890; nickname=entry.9876543210`.

Field names: `fullname`, `firstname`, `lastname`, `title`, `nickname`,
`tel`, `lineid`, `linedisplay`, `email`, `year`. Use `skip` to force a
question blank, or `col:C` to pull straight from a spreadsheet column.

To find an entry id, open the pre-filled link and read it from the URL.

---

## Short links

`shortenFormLinks()` fills a `short_url` column in the `Forms` tab — one
short link per form — so you never open a shortener website by hand.

```js
shortenFormLinks()
```

Run it from the editor after adding a form. It:

- builds each link itself (`?form=<key>&openExternalBrowser=1`), so the key
  is always right
- **follows the short link and checks it really lands on your form** before
  writing it to the sheet; a dead service or a wrong redirect is reported,
  never saved
- remembers what it made, so re-running does not create duplicates or burn
  through rate limits

By default it uses free keyless services (is.gd, v.gd, TinyURL), trying the
next one if a service is down.

### Using your own bit.ly account

For branded links and click statistics, put a bit.ly API token in
**Project Settings → Script Properties** as `BITLY_TOKEN`. It will be used
in preference to the free services. If the token is rejected, the log says
so rather than quietly falling back.

> Short links depend on someone else's service staying up. For anything
> printed, point the QR code at the full `/exec` link instead — nobody types
> a QR code, and it can never expire.

---

## Testing it before you send it out

You do **not** need a second student account. Only the identity check
(*"who is visiting?"*) depends on being signed in as someone else —
everything after that is data lookup, and can be dry-run from the editor.

### Check one student — `previewAs(email, formKey)`

```js
previewAs('6512345630@docchula.com', 'Test1')
```

Run it from the Apps Script editor (Run → `previewAs`, then View → Logs).
It prints the data read for that student, what each question receives, and
a pre-filled URL you can open to see the real form. Works for **any** row
on the roster, signed in as yourself.

### Check the whole class — `auditAllStudents(formKey)`

```js
auditAllStudents('Test1')
```

Dry-runs every roster row against one form and reports:

- rows with no email address, which can never be matched
- duplicate addresses, where the second row is unreachable
- two addresses sharing the part before `@` — matching accepts either, so
  the wrong student could be filled in
- how many students would get each question left blank, with examples

This is the one to run after adding a form or re-syncing the roster.

### The one test that does need another person

To confirm identity detection end-to-end, add **one** helper's Google
address to a spare `SYNCDATA` row (any Gmail works — it does not have to be
`@docchula.com`) and have them open the link. If they get *their* row's
data and not yours, the deployment settings are right. Remove the row
afterwards.

---

## When something does not fill in

**Add `&debug=1` to the link.** Instead of redirecting, the page reports the
email it detected, the roster row it matched, and a question-by-question
table of what was sent and what was dropped.

```
https://script.google.com/macros/s/<DEPLOYMENT_ID>/exec?form=Test2&debug=1
```

You can also run **`validateSetup()`** from the Apps Script editor to check
every configured form at once (Run → validateSetup, then View → Logs). It
flags missing keys, duplicate keys, links with no `?entry.…`, and forms
where nothing would be filled in.

Common causes:

| Symptom | Cause |
| ------- | ----- |
| Every student gets a blank form | Deployment is not **Execute as: User accessing the web app**, or the link was opened in an in-app browser that is signed out. |
| One form works, another does not | The second form's `template_url` is the plain link, not a pre-filled one — or its questions use marker words the script does not know. `&debug=1` shows which. |
| A field is blank for some students only | Those rows are empty in `SYNCDATA`. |
| "Unknown form key" | The `?form=` value does not match column A. The page lists the keys that do exist. |

---

## หมายเหตุ (ภาษาไทย)

- เพิ่มฟอร์มใหม่ = เพิ่ม **1 แถว** ในแท็บ `Forms` (คอลัมน์ A = ชื่อคีย์,
  B = ลิงก์ pre-filled, C = ลิงก์ฟอร์มเปล่า) แล้วส่งลิงก์
  `...exec?form=<ชื่อคีย์>&openExternalBrowser=1` — **ไม่ต้องแก้โค้ด**
- ลิงก์ในคอลัมน์ B ต้องมาจาก **⋮ → รับลิงก์ที่กรอกไว้ล่วงหน้า**
  (ต้องมี `?entry.…` ในลิงก์) ถ้าใช้ลิงก์ฟอร์มธรรมดาจะกรอกข้อมูลไม่ได้
- ต้อง Deploy แบบ **Execute as: User accessing the web app**
  ไม่เช่นนั้นสคริปต์จะไม่รู้ว่าใครเป็นผู้เข้าใช้ และทุกคนจะได้ฟอร์มเปล่า
- ถ้าข้อมูลไม่ขึ้น ให้เติม `&debug=1` ท้ายลิงก์ เพื่อดูว่าติดตรงไหน
- รหัสนิสิตจะถูกเว้นว่างไว้ให้กรอกเองเสมอ
