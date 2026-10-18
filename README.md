# حجز مواعيد طلبات الدمج

تطبيق Flask مستقل بقاعدة SQLite (ملف واحد `booking.db`).

## النشر على Railway

1. **New Project → Deploy from GitHub repo** واختر هذا المستودع.
2. **قاعدة البيانات الدائمة**: اربط Postgres بمتغير `DATABASE_URL`
   (مثال: `${{Postgres.DATABASE_URL}}`). التطبيق ينشئ جداوله في schema باسم `booking`.
   بدون `DATABASE_URL` يستخدم ملف SQLite، وعندها **يجب** إرفاق Volume على `/data`
   وإلا تُمسح كل الطلبات مع كل إعادة نشر.
3. من **Variables** أضف:
   - `ADMIN_PASSWORD` — كلمة مرور لوحة الإدارة (طويلة).
   - `SECRET_KEY` — نص عشوائي طويل (اختياري مع Postgres).
4. من **Settings → Networking → Custom Domain** أضف `sid.edutech-egy.com`،
   ثم أنشئ سجل `CNAME` باسم `sid` عند مزوّد الدومين بالقيمة التي يعرضها Railway.

أمر التشغيل وفحص السلامة (`/healthz`) مضبوطان في `railway.json`.

## التشغيل محليًا

```bash
pip install -r requirements.txt
cp .env.example .env   # ثم اكتب ADMIN_PASSWORD
python app.py          # http://127.0.0.1:5050
```

## الصفحات

| المسار | الوظيفة |
| --- | --- |
| `/` | فورم طلب الدمج |
| `/lookup` | استعلام عن موعد بالرقم القومي والهاتف |
| `/ticket/...` | إيصال الحجز (طباعة A4 بباركود) |
| `/admin` | الإعدادات، الأيام، الطلبات، التصدير إلى Excel |

## النسخ الاحتياطي

صدّر Excel لكل الأيام من لوحة الإدارة بشكل دوري.
لاسترجاع نسخة JSON: `python restore_backup.py backup.json` (يضيف الطلبات غير الموجودة فقط).
