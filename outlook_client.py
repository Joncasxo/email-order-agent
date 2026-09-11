# -*- coding: utf-8 -*-
"""Skaitymas iš klasikinio Outlook per COM. Veikia be administratoriaus teisių.

COM objektai naudojami tik pagalbinėse funkcijose (_..._com), kad jų nuorodos
būtų atlaisvintos PRIEŠ CoUninitialize — kitaip COM objektai atlaisvinami jau
uždarius COM aplinką ir procesas gali nulūžti.
"""
import datetime
import os
import tempfile

import pythoncom
import win32com.client

OL_MAIL_ITEM = 43        # olMail
OL_FOLDER_INBOX = 6      # olFolderInbox
MAX_ATTACHMENT_MB = 20   # didesni priedai praleidžiami
SKIP_EXT = {".p7s", ".ics", ".vcf", ".asc", ".sig"}


def _smtp_address(msg) -> str:
    """SMTP adresas — Exchange siuntėjams SenderEmailAddress būna /O=... formos."""
    try:
        if getattr(msg, "SenderEmailType", "") == "EX":
            sender = msg.Sender
            if sender is not None:
                exu = sender.GetExchangeUser()
                if exu is not None and exu.PrimarySmtpAddress:
                    return exu.PrimarySmtpAddress
        return msg.SenderEmailAddress or ""
    except Exception:
        return ""


def _sender_matches(sender_name: str, sender_email: str, filters) -> bool:
    """Filtro įrašas gali turėti kelis žodžius — visi jie turi būti siuntėjo
    eilutėje (bet kokia tvarka, todėl tinka ir „Pavardenis Vardenis")."""
    if not filters:
        return True
    hay = f"{sender_name} <{sender_email}>".lower()
    for entry in filters:
        words = [w for w in str(entry).lower().split() if w]
        if words and all(w in hay for w in words):
            return True
    return False


def _received_naive(item) -> datetime.datetime:
    rt = item.ReceivedTime
    return datetime.datetime(rt.year, rt.month, rt.day, rt.hour, rt.minute, rt.second)


PR_ATTACHMENT_HIDDEN = "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B"
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"


def _attachment_inline(att, html_lower: str) -> bool:
    """Ar priedas įterptas laiško tekste/paraše (logotipai, banneriai)?"""
    try:
        pa = att.PropertyAccessor
    except Exception:
        return False
    try:
        if pa.GetProperty(PR_ATTACHMENT_HIDDEN):
            return True
    except Exception:
        pass
    try:
        cid = pa.GetProperty(PR_ATTACH_CONTENT_ID)
        if cid and html_lower and f"cid:{str(cid).lower()}" in html_lower:
            return True
    except Exception:
        pass
    return False


def _read_attachments(item):
    """Grąžina [(failo_vardas, baitai, inline)], praleidžiant per didelius priedus.
    inline=True — priedas įterptas laiško tekste/paraše (dažniausiai logotipai)."""
    result = []
    try:
        count = item.Attachments.Count
    except Exception:
        return result
    html_lower = ""
    try:
        html_lower = (item.HTMLBody or "")[:200000].lower()
    except Exception:
        pass
    tmpdir = tempfile.mkdtemp(prefix="uzsak_att_")
    for i in range(1, count + 1):
        try:
            att = item.Attachments.Item(i)
            name = att.FileName or f"priedas_{i}"
            ext = os.path.splitext(name)[1].lower()
            if ext in SKIP_EXT:
                continue
            inline = _attachment_inline(att, html_lower)
            path = os.path.join(tmpdir, f"{i}_{name}"[:150])
            att.SaveAsFile(path)
            size = os.path.getsize(path)
            if size > MAX_ATTACHMENT_MB * 1024 * 1024:
                os.remove(path)
                continue
            with open(path, "rb") as f:
                data = f.read()
            os.remove(path)
            result.append((name, data, inline))
        except Exception:
            continue
    try:
        os.rmdir(tmpdir)
    except OSError:
        pass
    return result


def _notify_skip(on_skip, item, err):
    """Praneša apie laišką, kurio nepavyko perskaityti (pvz., dar sinchronizuojamas).
    Toks laiškas nepažymimas apdorotu, tad kita patikra (su 30 min užlaida) jį pakartos."""
    if on_skip is None:
        return
    try:
        subj = item.Subject or "(be temos)"
    except Exception:
        subj = "(nepavyko perskaityti temos)"
    try:
        on_skip(subj, str(err))
    except Exception:
        pass


def _scan_folder(folder, since, max_count, skip, sender_filters, store_id,
                 on_skip=None):
    """Peržiūri vieną aplanką (naujausi pirmi) ir grąžina tinkamus laiškus."""
    emails = []
    try:
        items = folder.Items
        items.Sort("[ReceivedTime]", True)  # naujausi viršuje
    except Exception:
        return emails
    scanned = 0
    for item in items:
        scanned += 1
        if scanned > 2000:  # saugiklis labai didelėms pašto dėžutėms
            break
        try:
            if getattr(item, "Class", None) != OL_MAIL_ITEM:
                continue
            received = _received_naive(item)
        except Exception as e:
            _notify_skip(on_skip, item, e)
            continue
        if received <= since:
            break
        try:
            sender_name = getattr(item, "SenderName", "") or ""
            sender_email = _smtp_address(item)
        except Exception as e:
            _notify_skip(on_skip, item, e)
            continue
        if not _sender_matches(sender_name, sender_email, sender_filters):
            continue
        if skip is not None and skip(item.EntryID):
            continue
        body = ""
        try:
            body = (item.Body or "")[:30000]
        except Exception:
            pass
        conversation_id = ""
        try:
            conversation_id = item.ConversationID or ""
        except Exception:
            try:
                conversation_id = item.ConversationTopic or ""
            except Exception:
                pass
        emails.append({
            "entry_id": item.EntryID,
            "store_id": store_id,
            "conversation_id": conversation_id,
            "subject": (item.Subject or "")[:500],
            "sender_name": sender_name,
            "sender_email": sender_email,
            "received": received,
            "body": body,
            "attachments": _read_attachments(item),
        })
        if len(emails) >= max_count:
            break
    return emails


def _fetch_emails_com(since, max_count, skip, sender_filters, on_skip=None):
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    store_id = inbox.StoreID

    # Tikriname tik „Gauta": poaplankiuose taisyklės palieka KOPIJAS, o jų
    # identifikatoriai kitokie — skenuojant juos tas pats užsakymas dubliuotųsi.
    return _scan_folder(inbox, since, max_count, skip, sender_filters, store_id,
                        on_skip=on_skip)


def fetch_emails_since(since: datetime.datetime, max_count: int = 50,
                       skip=None, sender_filters=None, on_skip=None):
    """Grąžina laiškus, gautus po `since` (naujausi pirmi).

    `skip` — funkcija skip(entry_id) -> bool: jei True, laiškas praleidžiamas
    dar prieš siunčiant jo priedus (jau apdorotiems laiškams sutaupo daug laiko).
    `sender_filters` — sąrašas filtrų; laiškas imamas tik jei siuntėjas atitinka
    bent vieną (tuščias sąrašas = imami visi).
    `on_skip(subject, klaida)` — kviečiama, kai laiško nepavyksta perskaityti
    (jis bus pakartotas kitą patikrą, bet apie tai lieka žyma žurnale)."""
    pythoncom.CoInitialize()
    try:
        return _fetch_emails_com(since, max_count, skip, sender_filters, on_skip)
    finally:
        pythoncom.CoUninitialize()


def _list_emails_com(since, max_count, sender_filters):
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    inbox = ns.GetDefaultFolder(OL_FOLDER_INBOX)
    store_id = inbox.StoreID
    try:
        items = inbox.Items
        items.Sort("[ReceivedTime]", True)
    except Exception:
        return []
    out = []
    scanned = 0
    for item in items:
        scanned += 1
        if scanned > 3000:
            break
        try:
            if getattr(item, "Class", None) != OL_MAIL_ITEM:
                continue
            received = _received_naive(item)
        except Exception:
            continue
        if received <= since:
            break
        try:
            sender_name = getattr(item, "SenderName", "") or ""
            sender_email = _smtp_address(item)
            subject = (item.Subject or "")[:500]
            entry_id = item.EntryID
        except Exception:
            continue
        if sender_filters and not _sender_matches(sender_name, sender_email, sender_filters):
            continue
        try:
            n_att = item.Attachments.Count
        except Exception:
            n_att = 0
        out.append({
            "entry_id": entry_id, "store_id": store_id, "subject": subject,
            "sender_name": sender_name, "sender_email": sender_email,
            "received": received, "n_att": n_att,
        })
        if len(out) >= max_count:
            break
    return out


def list_emails_since(since: datetime.datetime, max_count: int = 400,
                      sender_filters=None):
    """LENGVAS sąrašas peržiūrai: be laiško teksto ir be priedų skaitymo.

    Skirtas rankiniam laiškų pasirinkimui — priedus (lėčiausią dalį) skaitome
    tik tiems laiškams, kuriuos vartotojas tikrai pažymi analizei."""
    pythoncom.CoInitialize()
    try:
        return _list_emails_com(since, max_count, sender_filters)
    finally:
        pythoncom.CoUninitialize()


def _load_email_com(entry_id, store_id):
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    try:
        item = ns.GetItemFromID(entry_id, store_id)
    except Exception:
        # Laiskas galejo buti perkeltas i kita aplanka (taisykle) - tada EntryID
        # nebeatitinka to paties store; bandome be store nurodymo.
        item = ns.GetItemFromID(entry_id)
    body = ""
    try:
        body = (item.Body or "")[:30000]
    except Exception:
        pass
    conversation_id = ""
    try:
        conversation_id = item.ConversationID or ""
    except Exception:
        try:
            conversation_id = item.ConversationTopic or ""
        except Exception:
            pass
    return {
        "entry_id": entry_id, "store_id": store_id,
        "conversation_id": conversation_id,
        "subject": (item.Subject or "")[:500],
        "sender_name": getattr(item, "SenderName", "") or "",
        "sender_email": _smtp_address(item),
        "received": _received_naive(item),
        "body": body,
        "attachments": _read_attachments(item),
    }


def load_email(entry_id: str, store_id: str):
    """Pilnas vieno laiško paruošimas analizei (tekstas + priedai)."""
    pythoncom.CoInitialize()
    try:
        return _load_email_com(entry_id, store_id)
    finally:
        pythoncom.CoUninitialize()


def _get_attachments_com(entry_id, store_id):
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    item = ns.GetItemFromID(entry_id, store_id)
    return _read_attachments(item)


def get_attachments(entry_id: str, store_id: str):
    """Iš naujo paima laiško priedus (pvz., įkėlimui į Excel failą)."""
    pythoncom.CoInitialize()
    try:
        return _get_attachments_com(entry_id, store_id)
    finally:
        pythoncom.CoUninitialize()


def _create_draft_com(subject, body, attachment_path, display):
    outlook = win32com.client.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # olMailItem
    mail.Subject = subject
    mail.Body = body
    if attachment_path:
        mail.Attachments.Add(attachment_path)
    mail.Save()  # TIK juodrastis - siuncia pats vartotojas
    if display:
        mail.Display()  # atidaro langa: irasyti gaveja ir siusti


def create_draft_with_attachment(subject: str, body: str, attachment_path: str = None,
                                 display: bool = False):
    """Sukuria Outlook JUODRASTI su priedu (nesiuncia - tai daro vartotojas).
    display=True — iškart atidaro laiško langą."""
    pythoncom.CoInitialize()
    try:
        _create_draft_com(subject, body, attachment_path, display)
    finally:
        pythoncom.CoUninitialize()


def _open_email_com(entry_id, store_id):
    outlook = win32com.client.Dispatch("Outlook.Application")
    ns = outlook.GetNamespace("MAPI")
    item = ns.GetItemFromID(entry_id, store_id)
    item.Display()


def open_email(entry_id: str, store_id: str):
    """Atidaro laišką Outlook lange."""
    pythoncom.CoInitialize()
    try:
        _open_email_com(entry_id, store_id)
    finally:
        pythoncom.CoUninitialize()
