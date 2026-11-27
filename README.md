# Local Office AI — دستیار هوشمند تحلیل داده و اسناد

---

## ۰. وضعیت انتشار در GitHub (baseline برای بازبینی فنی)

سورس پایه ابتدا از بستهٔ `Local-Office-AI_UI_UX_Refined.zip` وارد شاخهٔ `main` شد
(مبنای ممیزی: `6df625659c97c0b7da1f78f33fa5fc37864a7580`). گزارش Phase 0 و آماده‌سازی
مرحله بعد در شاخهٔ `fix/multisource-extraction-integrity` نگهداری می‌شود. شاخهٔ شروع
مهندسی Prompt 01A برابر `feature/prompt-01a-data-model-metadata` است؛ `main` نباید
مستقیماً تغییر کند. این مخزن خصوصی است و فایل‌های اداری واقعی در آن قرار نمی‌گیرند.

### منبع و دامنه

- **منبع:** `Local-Office-AI_UI_UX_Refined.zip` (ارسال‌شده توسط مالک پروژه).
- **نوع بسته:** کد کامل برنامه (موتور + رابط) همراه با پاس بازطراحی رابط PySide6.
  طبق `docs/ui_ux_refined_report.md` تغییرات این پاس محدود به لایهٔ رابط است و موتور تحلیل،
  محاسبات `Decimal` و ساخت گزارش Excel دست‌نخورده مانده است.
- **آنچه در مخزن منتشر شده:** درخت واقعی سورس با همان ساختار بسته (بدون تغییر محتوا)،
  به‌همراه `.gitkeep` برای پوشه‌های خالی `data/` و `.gitignore` توسعه‌یافته.

### ساختار دایرکتوری

```
run_engine.py                 نقطهٔ ورود موتور (CLI)
run_ui.py                     نقطهٔ ورود رابط مدرن PySide6
Run_Local_Office_AI*.bat      لانچرهای ویندوز (CLI، GUI قدیم، مدرن، تست)
Deploy_Update.bat / DEPLOY.md  مانیفست استقرار فقط‌کد
src/local_office_ai/
  core/          تنظیمات، لاگ، استثناها، مدل‌های داده
  excel/         خواندن، تشخیص ساختار، نرمال‌سازی، آمار، گزارش‌نویسی
  ai/            کلاینت Ollama، ساخت Context، Prompt، اعتبارسنجی پاسخ، Image Studio
  ingestion/     خواندن چندفرمتی (Word/PowerPoint/PDF/TXT/تصویر)
  monitoring/    نمونه‌گیری psutil و شمارنده‌های GPU
  app/           مراحل Workflow و Coordinator
  ui/            CLI، wizard قدیم Tkinter، event bridge
  ui/desktop/    رابط اصلی مدرن (صفحات، تم RTL، فونت، Worker، پنل سخت‌افزار، ویجت شناور)
tests/                        مجموعهٔ pytest (تعداد آیتم‌ها وابسته به محیط؛ در CI فعلی پس از افزودن ۴ تست آماده‌سازی ۲۲۰ آیتم جمع‌آوری شد)
config/                       تنظیمات پیش‌فرض (settings.json، synonyms.json، ui_settings.json)
docs/                         گزارش‌ها و راهنماهای فنی همین نسخه
legacy/                       اسکریپت V4.2 دست‌نخورده و لانچر آن
data/{input,output,processed,temp}  پوشه‌های داده (خالی، فقط .gitkeep)
logs/                         محل لاگ در زمان اجرا (در مخزن خالی است)
```

### چه چیزی در بسته هست و چه چیزی نیست

**هست:** کل سورس `src/`، کل `tests/`، لانچرها، `requirements.txt`،
`requirements-documents.txt`، `requirements-dev.txt`، `pytest.ini`، `config/`، `docs/`، `legacy/`، `AGENTS.md`،
`DEPLOY.md`، `Deploy_Update.bat`.

**نیست (و عمداً جعل نشده):**

| مورد | وضعیت | اثر |
|---|---|---|
| `data/input/Book2_AI_V4_2_Monitored.xlsx` | در بستهٔ ZIP/مخزن نیست | ۳۸ تست با نشان `requires_sample` به آن وابسته‌اند و در CI با marker از gate عمومی کنار گذاشته می‌شوند |
| `data/input/General_Office_Demo.xlsx` | در بستهٔ ZIP نبود | فقط در مستندات به آن اشاره شده است |
| `reference/v5_sample_report/` | در بستهٔ ZIP نبود | فقط مرجع طراحی است و در زمان اجرا خوانده نمی‌شود |
| `venv/`، لاگ‌ها، خروجی‌های تولیدشده | به‌عمد حذف شده | با `.gitignore` کنار گذاشته شده‌اند |
| `pyproject.toml` / `setup.py` | در بسته نبود | بستهٔ نصب ندارد؛ ورودی‌ها از طریق `sys.path` در لانچرها و `conftest.py` کار می‌کنند |

هیچ فایل گمشده‌ای بازسازی یا حدس زده نشده است.

### نتیجهٔ تاریخی آزمون بستهٔ اولیه

> این نتیجه مربوط به آزمون بستهٔ اولیه، پیش از چهار تست آماده‌سازی فعلی است؛ برای وضعیت تازه
> باید `docs/quality_evidence.md` و آخرین اجرای GitHub Actions خوانده شود.

اجرا روی Linux، Python 3.13.15، PySide6 6.12.0، openpyxl 3.1.5، psutil 7.2.2،
pytest 9.2.0 با `QT_QPA_PLATFORM=offscreen`:

```
python -m pytest tests -o addopts="" -q
→ 216 tests: 174 passed, 4 skipped, 37 errors, 1 failure
```

- ۳۸ تست `requires_sample` به فایل خصوصی
  `data/input/Book2_AI_V4_2_Monitored.xlsx` وابسته‌اند. این فایل در مخزن قرار نمی‌گیرد؛
  CI از نشان `-m "not requires_sample"` استفاده می‌کند تا این تست‌ها را اجرا نکند.
- **۴ مورد skip** مربوط به نبود کتابخانه‌های اختیاری `docx` و `pdfplumber`
  (`requirements-documents.txt`) است.
- `tests/test_modern_ui.py` با نصب PySide6 اجرا شد: ۴ تست موفق، ۲ تست فقط به‌دلیل
  همان فایل نمونه error دادند.
- تست‌های ویندوزی (پنل شناور always-on-top، System Tray، GPU از `nvidia-smi`/LibreHardwareMonitor)
  روی Linux قابل تأیید نیستند و در گزارش‌های `docs/` هم به‌عنوان تأییدنشده ثبت شده‌اند.

### شروع کار برای AI بعدی — Prompt 01A

Phase 0 قبلاً انجام و مستند شده است. برای شروع، `docs/NEXT_AGENT_START_HERE.md`
را دنبال کنید. شاخهٔ شروع `feature/prompt-01a-data-model-metadata` است.

عامل بعدی باید مستندات موجود را بخواند و فقط branch/HEAD/working tree را تأیید کند؛
نیازی به اجرای مجدد ممیزی Phase 0 یا درخواست دوبارهٔ handover/audit ندارد. سپس متن
اصلی Prompt 01A را که مالک پروژه ارائه می‌کند اجرا کند. Prompt 01A هنوز اجرا نشده است.

برای اجرای gate عمومی تست‌ها (بدون فایل مرجع محرمانه):

```bash
python -m pytest tests -o addopts="" -q -ra -m "not requires_sample"
```

### محدودیت‌های شناخته‌شده

- فایل نمونهٔ مرجع و سه فایل اداری واقعی در مخزن نیستند؛ ۳۸ تست `requires_sample`
  فقط در محیط محلی مجاز و با فایل نمونه قابل اجرا هستند.
- بخشی از مستندات به مسیر ویندوزی `F:\Local-Office-AI` و `venv\Scripts\python.exe` وابسته است.
- فایل `Local-Office-AI_UI_UX_Refined.zip` در ریشهٔ مخزن باقی مانده است (از کامیت قبلی `main`)
  و حذف آن نیاز به تأیید مالک دارد.

---


## ورودی چندفرمتی و تحلیل ترکیبی

نسخه فعلی از انتخاب هم‌زمان چند فایل و افزودن پوشه/زیرپوشه‌ها به یک تحلیل مشترک پشتیبانی می‌کند. فرمت‌ها شامل XLSX، XLSM، CSV، DOCX، PPTX، PDF، TXT و تصویرهای رایج هستند. وابستگی‌های خواندن اسناد در `requirements-documents.txt` جدا شده‌اند تا نصب هسته اصلی تکرار نشود. راهنمای کامل، محدودیت فرمت‌های قدیمی و پیش‌نیاز OCR فارسی در `docs/MULTIFORMAT_INPUTS.md` قرار دارد.


بازطراحی و توسعهٔ `Ollama_Excel_Assistant_V4_2.py` به یک موتور عمومی، ماژولار و تست‌شده
برای **پردازش، تحلیل آماری و گزارش‌گیری داده‌های اداری و سازمانی** است. بستهٔ اول پایهٔ
معماری است و قابلیت‌های منتخب بستهٔ V5 (واحدهای مقیاس، مترادف نام مؤسسات، تشخیص معنایی
ترکیبی و اعتبارسنجی دقیق اعداد) با همان قراردادهای داده در آن ادغام شده‌اند.

موتور فایل‌های Excel را می‌خواند، ساختار و نوع واقعی ستون‌ها را از عنوان و محتوا تشخیص
می‌دهد، داده‌ها را با `Decimal` استاندارد می‌کند، آمار توصیفی و توزیع فراوانی می‌سازد،
گروه‌ها را مقایسه می‌کند، جدول متقاطع، روند زمانی، همبستگی پیرسون/اسپیرمن و پرت‌های IQR را
محاسبه می‌کند و کیفیت داده (خالی، تکراری، ثابت، مختلط و شناسه) را گزارش می‌دهد. تحلیل
وام/تعهد یک **ماژول اختیاری** برای فایل‌هایی است که واقعاً چنین ستون‌هایی دارند؛ Book2 فقط
نمونهٔ نمایشی این ماژول است، نه محور محصول.

> فایل ورودی هرگز بازنویسی نمی‌شود. خروجی همیشه در مسیر جداگانه ساخته می‌شود و هش فایل
> ورودی پیش و پس از پردازش مقایسه می‌شود.

---

## ۱. نصب

محیط هدف از قبل آماده است: `F:\Local-Office-AI\venv` با Python 3.14، PySide6 6.12 و
کتابخانه‌های `openpyxl`، `psutil` و `XlsxWriter`. **نیازی به نصب مجدد Python، ساخت venv
تازه یا نصب دوبارهٔ وابستگی‌ها نیست.** فقط اگر `venv` بازسازی شد:

```bat
cd /d F:\Local-Office-AI
venv\Scripts\python.exe -m pip install -r requirements.txt
```

وابستگی‌های مخصوص تست (`pytest` و `reportlab`) در `requirements-dev.txt` جدا شده‌اند و
برای اجرای برنامه لازم نیستند:

```bat
venv\Scripts\python.exe -m pip install -r requirements-dev.txt
```

وابستگی‌های موتور `openpyxl` و `psutil` هستند و رابط مدرن به `PySide6` نیاز دارد.
`pandas`/`numpy` لازم نیست؛ محاسبات با `decimal` و `fractions` انجام می‌شوند. Ollama (اختیاری برای تحلیل متنی) و ComfyUI
(اختیاری برای تولید تصویر) از قبل نصب‌شده فرض می‌شوند و این بسته آن‌ها را نصب یا دانلود
نمی‌کند.

---

## ۲. اجرا

در رابط گرافیکی، فیلد **درخواست (فارسی)** همان کاری را می‌کند که `--request` در خط فرمان؛
دکمهٔ **تحلیل پوشه** هم فایل‌های انتخاب‌شدهٔ یک پوشه را تحلیل و در یک گزارش مقایسه می‌کند.

```bat
:: ویندوز — رابط خط فرمان (لانچر اصلی، فقط venv پروژه)
Run_Local_Office_AI.bat --file "F:\Local-Office-AI\data\input\Book2_AI_V4_2_Monitored.xlsx"

:: ویندوز — رابط مدرن PySide6 (توصیه‌شده)
Run_Local_Office_AI_Modern.bat

:: ویندوز — فقط ویجت شناور سخت‌افزار، بدون پنجرهٔ اصلی
Run_Local_Office_AI_Modern.bat --floating-only

:: ویندوز — پنجرهٔ اصلی همراه ویجت شناور
Run_Local_Office_AI_Modern.bat --floating

:: ویندوز — رابط قدیمی Tkinter (مسیر بازگشت)
Run_Local_Office_AI_GUI.bat

:: ویندوز — تست‌ها
Run_Tests.bat

:: اجرای مستقیم با Python همان venv
"F:\Local-Office-AI\venv\Scripts\python.exe" run_engine.py ^
  --file "F:\Local-Office-AI\data\input\Book2_AI_V4_2_Monitored.xlsx" ^
  --model qwen3.5:9b
```

هیچ لانچری از `py` سراسری یا Python سیستم استفاده نمی‌کند. فایل قدیمی
`Run_Ollama_Excel_Assistant_V4_2.bat` فقط در `legacy/` نگهداری می‌شود و لانچر اصلی نیست.

گزینه‌های مهم:

| گزینه | کار |
|---|---|
| `--file`, `-f` | فایل Excel ورودی (`.xlsx` / `.xlsm`) |
| `--request`, `-r` | درخواست کاربر به فارسی؛ نوع تحلیل، شاخص‌ها، ماه‌ها و سال‌ها از همین متن خوانده می‌شود |
| `--folder` | تحلیل همهٔ فایل‌های Excel یک پوشه و مقایسهٔ بین آن‌ها |
| `--audit` | افزودن برگه‌های فنی و ردیابی منبع (حالت توسعه/حسابرسی) |
| `--model`, `-m` | مدل Ollama؛ پیش‌فرض انتخاب خودکار مدل Qwen |
| `--no-ai` | تحلیل مدل را غیرفعال و گزارش را با قواعد قطعی Python بساز |
| `--task` | `general_analysis` \| `data_quality` \| `trend_analysis` \| `group_comparison` \| `correlation_analysis` \| `management_analysis` \| … |
| `--question` | پرسش کاربر دربارهٔ فایل |
| `--goal` | هدف تحلیل به زبان طبیعی؛ در GUI هم فیلد جداگانه دارد |
| `--output`, `-o` | مسیر خروجی |
| `--json` | ذخیرهٔ خلاصهٔ ماشین‌خوان |
| `--check-structure` | فقط تشخیص ساختار؛ بدون تولید فایل |
| `--list-models` | فهرست مدل‌های نصب‌شده Ollama |
| `--list-checkpoints` | فهرست checkpoint های ComfyUI |
| `--image [PROMPT]` | تولید تصویر با ComfyUI پس از تحلیل |
| `--yes`, `-y` | تأیید خودکار ساختار کم‌اطمینان (حالت غیرتعاملی) |
| `--gui` | رابط گرافیکی Tkinter |
| `--root` | ریشهٔ پروژه (پیش‌فرض: محل همین بسته) |
| `--quiet` | خروجی کنسول کم‌تر |

اگر ساختار فایل با اطمینان کمتر از ۶۰٪ تشخیص داده شود، در حالت تعاملی از کاربر تأیید
گرفته می‌شود؛ در حالت غیرتعاملی (سرویس/تست) پردازش با ثبت هشدار ادامه می‌یابد.

---

## ۲.۱. رابط مدرن PySide6

رابط جدید کاملاً فارسی و راست‌به‌چپ است و در `src/local_office_ai/ui/desktop/` قرار دارد.
موتور، فایل‌های V4.2، اتصال Ollama و رابط Tkinter دست‌نخورده‌اند. امکانات اصلی:

- خانه، تحلیل فایل، تحلیل پوشه، گفت‌وگوی هوشمند، گزارش‌های قبلی و تنظیمات.
- انتخاب فایل یا پوشه، اعتبارسنجی واقعی Excel و Drag & Drop.
- مراحل واقعی موتور با Worker و بدون قفل‌شدن رابط؛ بدون درصد ساختگی.
- خلاصه، یافته‌ها، جدول، نمودار، هشدارها، پیش‌نمایش و بازکردن Excel خروجی.
- انتخاب فونت فارسی نصب‌شده، اندازهٔ متن/عنوان، پیش‌نمایش، فونت سفارشی محلی و
  ذخیرهٔ انتخاب‌ها در `config/ui_settings.json`.
- تم روشن/تیره، انتخاب مدل Ollama، مسیر ورودی/خروجی و تنظیمات پایش.
- پنل جمع‌شوندهٔ سخت‌افزار با `psutil`، `nvidia-smi`، `sysfs` و در صورت وجود
  LibreHardwareMonitor؛ مقدار ناموجود «در دسترس نیست» است و هرگز به Excel اضافه
  نمی‌شود.
- ویجت شناور مستقل سخت‌افزار: پنجرهٔ کوچک بدون قاب و همیشه روی سایر پنجره‌ها،
  جابه‌جایی با ماوس، شفافیت قابل تنظیم، سینی سیستم، ذخیرهٔ محل و اجرای اختیاری هنگام
  ورود به ویندوز. همان `HardwareMonitor` را استفاده می‌کند و فقط هنگام نمایش
  نمونه‌برداری می‌کند.
- جدول‌های راست‌به‌چپ با سرستون و سلول راست‌چین، اعداد طولانی کوتاه‌شده در نمایش،
  Tooltip برای متن‌های بلند و انتقال خودکار به نتیجه پس از پایان تحلیل.

راهنمای کامل: [`docs/MODERN_UI.md`](docs/MODERN_UI.md) — گزارش تحویل جاری:
[`docs/OPERATIONAL_DESKTOP_REPORT.md`](docs/OPERATIONAL_DESKTOP_REPORT.md)؛ گزارش
تاریخی نسخهٔ پایه: [`docs/MODERN_UI_REPORT.md`](docs/MODERN_UI_REPORT.md).

## ۳. خروجی

مسیر پیش‌فرض: `data/output/<نام فایل>_AI_Intelligence.xlsx` (در صورت وجود، پسوند شماره‌دار
می‌شود تا هیچ خروجی قبلی بازنویسی نشود). خروجی ماشین‌خوان در
`data/processed/<نام فایل>_AI_Intelligence.json` ذخیره می‌شود و شامل مقادیر خام هر سلول
(`raw_cells`) است. لاگ: `logs/local_office_ai.log` (چرخشی، UTF-8).

نمونهٔ اجرای نهایی و تأییدشدهٔ همین بسته در `reports/final_run/` نگه داشته شده است
(`Book2_Final_Validation_Report.xlsx` و JSON متناظر).

### خروجی تمیز (پیش‌فرض)

خروجی پیش‌فرض فقط چیزی را دارد که برای پاسخ به درخواست لازم است؛ هیچ برگهٔ پایش منابع،
گزارش فنی یا اطلاعات اجرای داخلی در آن نیست. برگه‌ها بر اساس درخواست و ساختار داده
انتخاب می‌شوند:

| برگه | چه زمانی ساخته می‌شود |
|---|---|
| پاسخ درخواست | درخواست کاربر، تفسیر آن، پاسخ‌های محاسبه‌شده با Python برای هر شاخص و منبع هر عدد |
| داشبورد | همیشه؛ خلاصهٔ شاخص‌های کلیدی و نمودار متناسب با داده |
| خلاصه اجرایی | همیشه؛ خلاصهٔ مدیریتی، یافته‌ها، اطلاعات ناقص و پیشنهادها |
| جدول ماهانه شاخص‌ها | اگر فایل ساختار ماهانه داشته باشد؛ جمع هر ماه از سلول‌های همان برگه |
| مقایسه دوره‌ای شاخص‌ها | اگر فایل دو دوره داشته باشد یا درخواست مقایسه باشد |
| شناسنامه داده | همیشه؛ پروفایل، پوشش، یکتایی و کیفیت هر ستون |
| آمار توصیفی | اگر شاخص عددی معنادار وجود داشته باشد |
| توزیع فراوانی / تحلیل گروه‌ها / روند و همبستگی | فقط اگر داده و درخواست پشتیبانی کند |
| اعتبارسنجی داده | همیشه؛ چک‌لیست ۹ بررسی با وضعیت قبول/هشدار/رد |
| موارد نیازمند بررسی | اگر ناهنجاری یا مغایرت واقعی پیدا شود |

اگر درخواست فقط یک جدول یا یک عدد بخواهد، همان ساخته می‌شود؛ برگهٔ اضافی تولید نمی‌شود.

وقتی مدل محلی فعال است، متن مدیریتی (خلاصه، یافته‌ها، پیشنهادها) داخل همین برگه‌های
پاسخ و خلاصه اجرایی می‌آید؛ برگهٔ جدا برای «اجرای هوش مصنوعی» ساخته نمی‌شود.

در تحلیل پوشه، برگهٔ «خلاصه پوشه» سه ردیف «درخواست کاربر»، «تفسیر درخواست» و
«پاسخ درخواست» دارد و مقایسهٔ شاخص‌های مشترک در برگهٔ «مقایسه بین فایل‌ها» است.

### پروفایل حسابرسی (`--audit`)

برای عیب‌یابی و بازبینی توسعه‌دهنده، `--audit` برگه‌های زیر را هم می‌نویسد: «داده خام»،
«داده استانداردشده»، «ردیابی منبع»، «گزارش عملکرد»، «نمونه‌های پایش»، «نمونه‌های GPU»،
«محدودیت‌ها و خطاها»، «ساختار فایل»، «تحلیل هوش مصنوعی»، «نتایج محاسبات»، «روش‌شناسی
آماری» و برگه‌های اختیاری وام/تعهد. این پروفایل هرگز پیش‌فرض نیست.

قالب‌بندی: راست‌چین (RTL)، فونت Tahoma، قالب عدد و تاریخ شمسی، رنگ‌بندی وضعیت، Freeze
Panes، AutoFilter، نمودار، و تنظیمات چاپ افقی Fit-to-Width.

---

## ۴. معماری

```text
src/local_office_ai/
├── core/        settings, logging_config, exceptions, models
├── excel/       reader, workbook_inspector, sheet_analyzer, table_detector,
│                data_extractor, normalization, calculations, statistical_analysis,
│                generic_report, reconciliation, anomaly_detection, report_writer
├── ai/          ollama_client, context_builder, prompt_manager,
│                response_validator, analyzer, image_studio
├── monitoring/  resource_monitor, performance_report
├── app/         workflow, coordinator
└── ui/          cli, wizard, event_bridge
```

`config/settings.json` تنظیمات و `config/synonyms.json` مترادف نام مؤسسات را نگه می‌دارد.
نمونهٔ گزارش مرجع V5 در `reference/v5_sample_report/` و کد قدیمی V4.2 در `legacy/` است.

قرارداد هر ماژول:

- `reader` → `WorkbookBundle` (دو نسخهٔ مقدار و فرمول، هش، هشدارها)
- `workbook_inspector` → `WorkbookMeta`/`SheetMeta` (محدودهٔ داده در برابر محدودهٔ فرمت)
- `sheet_analyzer` → `SheetGrid`, `HeaderCandidate`, پروفایل ستون و نقش معنایی
- `table_detector` → `TableCandidate` (نوع، اطمینان، دلیل تشخیص)
- `data_extractor` → `LoanRecord`/`ObligationRecord`/`DeclaredTotal` + ردیف‌های استانداردشده
- `calculations` → `Calculation` با روش، منابع و کنترل دوباره
- `statistical_analysis` → `GenericAnalysis` (پروفایل ستون، آمار توصیفی، فراوانی، گروه، جدول متقاطع، روند، همبستگی، پرت و کیفیت)
- `reconciliation` → `ReconciliationItem` با سه وضعیت مستقل
- `anomaly_detection` → `Anomaly` با شدت و مکان
- `report_writer` → فایل Excel (نوشتن اتمیک + بازخوانی تأییدی)
- `workflow` → ارکستراسیون مرحله‌به‌مرحله با زمان‌سنجی و انتساب خطا به مرحله
- `coordinator` → اتصال تنظیمات/لاگ/پایش و چاپ خلاصه

هیچ ماژولی به نام برگه، ستون یا فایل نمونه وابسته نیست.

---

## ۵. قواعد محاسبه و هوش مصنوعی

- Python تنها مرجع محاسبه است. همهٔ اعداد و مبالغ `Decimal` هستند؛ آماره‌ها با
  `Decimal`/`statistics` محاسبه می‌شوند و هیچ نتیجهٔ عددی از مدل زبانی گرفته نمی‌شود.
- تحلیل آماری اکتشافی است: میانگین، میانه، مد، واریانس نمونه (n-1) و جامعه (n)، انحراف
  معیار، چارک/صدک با درون‌یابی R-7، IQR، بازه‌های استرجس، پیرسون/اسپیرمن و پرت IQR.
  مقادیر خالی و پرت حذف نمی‌شوند؛ فقط علامت‌گذاری و شمارش می‌شوند.
- واحدهای مقیاس در متن («هزار»، «میلیون»، «میلیارد»، «همت») پیش از تبدیل اعمال می‌شوند؛
  «حدود ۲ میلیون تومان» به `2,000,000` تبدیل می‌شود، نه `2`. هر تبدیل در `warnings` ثبت
  می‌شود و مقدار همیشه `Decimal` می‌ماند.
- جمع‌های اعلامی فایل از سلول‌های برچسب‌دار («جمع کل»، «مجموع» …) خوانده و با محاسبهٔ
  مستقل مقایسه می‌شوند؛ نتیجه یکی از `verified` / `mismatch` / `not_verifiable` /
  `no_declared_value` است.
- مدل فقط توضیح می‌دهد: هر عدد در پاسخ مدل با مجموعهٔ اعداد تأییدشده مقایسه می‌شود؛ عدد
  بی‌منبع به‌عنوان `unsupported` ثبت و پاسخ «غیرمتکی» علامت می‌خورد. اگر مدل از مغایرت
  سخن بگوید ولی حسابرسی Python مغایرتی نیافته باشد، تناقض ثبت می‌شود.
- **سیاست مقایسهٔ اعداد مالی:** مقدار `Decimal` استاندارد، اختلاف مجاز **صفر**، بدون
  گردکردن، با نرمال‌سازی واحد پول و واحد مقیاس. تلورانس نسبی ۰٫۱٪ وجود ندارد؛
  `325,000,001` در برابر `325,000,000` رد می‌شود. این سیاست در برگهٔ «تحلیل هوش مصنوعی»
  و در خروجی JSON ثبت می‌شود (`NumberComparisonPolicy`).
- اگر Ollama در دسترس نباشد، گزارش با خلاصهٔ قطعی فارسی از همان اعداد ساخته می‌شود و
  دلیل دقیق در برگهٔ تحلیل و گزارش عملکرد ثبت می‌گردد. شکست مدل هرگز پردازش را متوقف نمی‌کند.
- حجم دادهٔ ارسالی به مدل با بودجهٔ کاراکتر محدود و روند کوتاه‌سازی ثبت می‌شود
  (`trimmed_rows`, `trimmed_anomalies`, `degraded`).

---

## ۶. پایش منابع (فقط عیب‌یابی اختیاری)

پایش منابع بخشی از پردازش عادی نیست: در پروفایل تمیز نه اندازه‌گیری می‌شود و نه در
خروجی می‌آید. این قابلیت فقط در پروفایل `--audit` و برای عیب‌یابی توسعه‌دهنده فعال است.

- CPU/RAM سیستم و مصرف پردازش‌های Python/Ollama با `psutil` (نمونه‌برداری در نخ جداگانه).
- GPU در ویندوز از `typeperf` با شمارنده‌های `GPU Engine` و `GPU Adapter Memory` خوانده
  می‌شود؛ انتساب VRAM به Ollama فقط وقتی انجام می‌شود که نام نمونه شامل PID همان پردازش
  باشد.
- مقدار اندازه‌گیری‌نشده `None` است و در گزارش «در دسترس نبود» نمایش داده می‌شود، نه صفر.
- درصد یک GPU Engine هرگز به‌عنوان درصد کل GPU معرفی نمی‌شود.
- محدودیت‌های ویندوز (نبود انتساب دقیق VRAM، محدودیت شمارنده‌ها، فاصلهٔ نمونه‌برداری) در
  برگهٔ «محدودیت‌ها و خطاها» نوشته می‌شوند.

---

## ۷. پیکربندی

`config/settings.json` (همهٔ مسیرها نسبت به ریشهٔ پروژه حل می‌شوند، پس روی `F:\` یا هر
مسیر دیگری یکسان کار می‌کند):

```json
{
  "data_dir": "data",
  "max_file_mb": 200,
  "max_processing_seconds": 1800,
  "ollama_base_url": "http://127.0.0.1:11434",
  "ollama_timeout_seconds": 300,
  "ollama_num_ctx": 8192,
  "ollama_retries": 2,
  "model_name_prefixes": ["qwen3.5", "qwen3", "qwen2.5"],
  "max_ai_context_chars": 9000,
  "max_ai_prompt_chars": 11000,
  "max_ai_sheets_per_chunk": 3,
  "ai_chunk_retry_ratio": 0.55,
  "monitor_interval_seconds": 1.0,
  "gpu_interval_seconds": 2.5
}
```

ریشهٔ پروژه با `LOCAL_OFFICE_AI_ROOT` یا `--root` قابل تغییر است. کلیدهای ناشناخته در
`settings.extra` نگه داشته می‌شوند و خطا نمی‌دهند.

`config/synonyms.json` فقط برای **تطبیق** نام مؤسسات است:

```json
{
  "_comment": "کلید باید زیررشتهٔ امن باشد؛ متن منبع تغییر نمی‌کند.",
  "صندق": "صندوق",
  "قرضالحسنه": "قرض الحسنه"
}
```

مترادف‌ها روی `institution_key` اعمال می‌شوند؛ `institution_raw` و مقدار سلول اصلی دست‌نخورده
می‌مانند. تطبیق احتمالی همچنان نیازمند تأیید انسانی است. برای استفاده از فایل دیگر، کلید
`synonyms_file` را در `settings.json` بگذارید (مسیر نسبی به ریشهٔ پروژه حل می‌شود).

---

## ۸. وضعیت واقعی قابلیت‌ها

| قابلیت | وضعیت | شاهد |
|---|---|---|
| انتخاب/خواندن فایل، عدم تغییر ورودی | IMPLEMENTED + TESTED | `test_reader_and_inspector`, `test_workflow_v42` |
| رد فایل خراب/قدیمی/بزرگ/قفل‌شده | IMPLEMENTED + TESTED | همان |
| تشخیص ابعاد داده در برابر محدودهٔ فرمت | IMPLEMENTED + TESTED | `A1:H15` در برابر `A1:BA511` |
| تشخیص چند جدول در یک برگه و کنار هم | IMPLEMENTED + TESTED | `test_sheet_structure` |
| تشخیص عنوان/سرستون/ردیف جمع/ردیف برچسب-مقدار | IMPLEMENTED + TESTED | `test_sheet_structure` |
| نقش معنایی ستون‌ها و نوع داده | IMPLEMENTED + TESTED | همان |
| تشخیص معنایی ترکیبی عنوان + محتوای واقعی و گزارش تناقض | IMPLEMENTED + TESTED | `test_units_synonyms_semantics`, برگهٔ «ساختار فایل» |
| سلول ادغام‌شده، مخفی، خطا، فرمول | IMPLEMENTED + TESTED | `test_reader_and_inspector`, `test_edge_cases` |
| استانداردسازی اعداد فارسی/عربی/جداکننده/منفی/متن | IMPLEMENTED + TESTED | `test_normalization` |
| واحدهای مقیاس هزار/میلیون/میلیارد/همت با Decimal | IMPLEMENTED + TESTED | `test_units_synonyms_semantics` |
| مترادف نام مؤسسات فقط برای تطبیق (بدون تغییر مقدار اصلی) | IMPLEMENTED + TESTED | `test_units_synonyms_semantics` |
| تاریخ شمسی/میلادی/سریال Excel و تاریخ مبهم | IMPLEMENTED + TESTED | `test_normalization` |
| یکسان‌سازی متن، نیم‌فاصله، کاراکتر نامرئی، صفر ابتدایی | IMPLEMENTED + TESTED | همان |
| محاسبات Decimal + کنترل دوباره | IMPLEMENTED + TESTED | `test_calculations` |
| حسابرسی جمع اعلامی و مغایرت | IMPLEMENTED + TESTED | `test_calculations`, `test_workflow_v42` |
| تطبیق قطعی/احتمالی/عدم تطبیق + تعهد خارج جدول | IMPLEMENTED + TESTED | `test_reconciliation_and_anomalies`, `test_units_synonyms_semantics` |
| تشخیص ناهنجاری (تکرار، منفی، خارج‌از‌محدوده، خطا، مخفی، محافظت) | IMPLEMENTED + TESTED | `test_reconciliation_and_anomalies`, `test_edge_cases` |
| چک‌لیست اعتبارسنجی داده (۹ بررسی، قبول/هشدار/رد) | IMPLEMENTED + TESTED | `test_validation` |
| حفظ مقدار خام هر سلول در خروجی و JSON پردازش‌شده | IMPLEMENTED + TESTED | `test_validation`, `test_workflow_v42` |
| مانده اقساط (اقساط باقی‌مانده و مبلغ باقی‌مانده) | IMPLEMENTED + TESTED | `test_validation` |
| گروه‌بندی و تجمیع عمومی جدول‌های غیروام + سهم درصدی | IMPLEMENTED + TESTED | `test_validation` |
| سقف زمان پردازش و توقف کنترل‌شده | IMPLEMENTED + TESTED | `test_validation` |
| شناسایی و انتخاب مدل Ollama | IMPLEMENTED + TESTED (سرور HTTP واقعی محلی) | `test_ai_layer` |
| ارسال دادهٔ خلاصه‌شده و محدودسازی Context | IMPLEMENTED + TESTED | `test_ai_layer` |
| اعتبارسنجی پاسخ مدل (عدد بی‌منبع/تناقض/JSON نامعتبر) | IMPLEMENTED + TESTED | `test_ai_layer`, `test_ai_exact_policy` |
| مقایسهٔ دقیق Decimal با اختلاف صفر (بدون تلورانس ۰٫۱٪) | IMPLEMENTED + TESTED | `test_ai_exact_policy`, `test_ai_layer` |
| گزارش فارسی بدون مدل (fallback) | IMPLEMENTED + TESTED | `test_ai_layer`, `test_workflow_v42` |
| تحلیل زندهٔ Qwen3.5 | IMPLEMENTED، اجرای واقعی BLOCKED | Ollama در این محیط نصب نیست |
| موتور آماری عمومی (پروفایل، توصیفی، فراوانی، گروه، متقاطع، روند، همبستگی، پرت، کیفیت) | IMPLEMENTED + TESTED | `test_generic_statistics` |
| انتخاب جدول اصلی بر اساس کیفیت داده، نه فرض وام | IMPLEMENTED + TESTED | `test_workflow_generic`, `test_workflow_v42` |
| گزارش خروجی عمومی با برگه‌های منتخب و بدون برگه‌های خالی وام | IMPLEMENTED + TESTED | `test_generic_statistics`, `test_validation` |
| خروجی تمیز پیش‌فرض بدون برگه‌های پایش/فنی | IMPLEMENTED + TESTED | `test_output_quality` |
| خروجی کامل حسابرسی با `--audit` | IMPLEMENTED + TESTED | `test_output_quality`, `test_report_writer` |
| پاسخ درخواست فارسی: تفسیر، پاسخ هر شاخص، منبع هر عدد | IMPLEMENTED + TESTED | `test_output_quality`, `test_generic_statistics` |
| جدول ماهانه و مقایسهٔ دوره‌ای از سلول‌های هر برگه | IMPLEMENTED + TESTED | `test_output_quality` |
| تحلیل پوشه و مقایسهٔ چند فایل بدون تغییر فایل‌های ورودی | IMPLEMENTED + TESTED | `test_output_quality` |
| انتقال درخواست فارسی به workflow از رابط گرافیکی | IMPLEMENTED + TESTED (بررسی AST) | `test_gui_bridge` |
| متن مدل در پروفایل تمیز، بدون برگهٔ اجرای هوش مصنوعی | IMPLEMENTED + TESTED | `test_output_quality` |
| تشخیص دوره/سال از سرستون، عنوان و نام برگه + گزارش تعارض | IMPLEMENTED + TESTED | `test_structure_intelligence`, `test_output_quality` |
| RTL، قالب عدد/تاریخ، رنگ وضعیت، فریز، فیلتر، نمودار، چاپ | IMPLEMENTED + TESTED | همان |
| نوشتن اتمیک + بازخوانی تأییدی خروجی | IMPLEMENTED + TESTED | همان |
| زمان‌سنجی مرحله‌ها و کل پردازش | IMPLEMENTED + TESTED | `test_workflow_v42`, `test_monitoring` |
| توکن ورودی/خروجی و سرعت تولید | IMPLEMENTED + TESTED (از فیلدهای API) | `test_ai_layer` |
| CPU/RAM سیستم و پردازش‌ها | IMPLEMENTED + TESTED | `test_monitoring` |
| GPU Engine / Dedicated VRAM / Shared GPU | IMPLEMENTED، TESTED فقط مسیر «در دسترس نبود» | ویندوز لازم است |
| ثبت محدودیت و خطای پایش | IMPLEMENTED + TESTED | `test_monitoring` |
| لاگ چرخشی UTF-8 | IMPLEMENTED + TESTED | `test_edge_cases` |
| تأیید ساختار کم‌اطمینان (CLI/GUI) | IMPLEMENTED + TESTED (callback) | `test_workflow_generic` |
| رابط گرافیکی Tkinter | IMPLEMENTED + TESTED (Queue/رویداد)، اجرای واقعی BLOCKED | `test_gui_bridge`؛ `libtk8.6.so` در این محیط نیست |
| رابط مدرن PySide6 (RTL/فونت/پایش) | IMPLEMENTED + TESTED (پنجره، موتور، خروجی، فونت، پایش) | `test_modern_ui`؛ 6 تست پذیرش بدون Mock |
| ویجت شناور سخت‌افزار (Always on Top/سینی/شفافیت) | IMPLEMENTED + TESTED (offscreen)؛ اجرای واقعی ویندوز BLOCKED | `test_floating_monitor`؛ بدون Mock |
| عدم دسترسی Thread پردازشی به Widgetها | IMPLEMENTED + TESTED (بررسی AST و bridge) | `test_gui_bridge` |
| لانچر ویندوز `.bat` با venv پروژه | IMPLEMENTED + TESTED (بررسی متن/مسیر)، اجرا BLOCKED | `test_packaging`؛ فقط ویندوز |
| استودیو تصویر ComfyUI (حفظ‌شده از V4.2) | IMPLEMENTED + TESTED (سرور HTTP واقعی محلی) | `test_image_studio` |
| مسیرهای درایو `F:` | IMPLEMENTED به‌صورت نسبی به ریشه، BLOCKED برای آزمون | درایو `F:` اینجا وجود ندارد |

---

## ۹. تست

```bash
python -m pytest tests -o addopts="" -q
```

آخرین اجرای کامل گزارش‌شده در بستهٔ قبلی **191 passed** بود (187 تست موتور + 4 تست
پذیرش رابط مدرن). در این بسته ۱۶ تست جدید اضافه شد و مجموعه به **۲۰۷ تست** رسید؛
طبق درخواست، همهٔ تست‌های قبلی موتور دوباره اجرا نشدند و فقط تست‌های جدید و فایل‌های تغییریافته
اجرا شدند: `test_deploy_manifest` (۵)، `test_floating_monitor` (۸)، دو تست در
`test_modern_ui` (ویجت و انتقال واقعی Ollama) و یک تست کیفیت پاسخ در `test_output_quality`.
`tests/test_output_quality.py` قرارداد خروجی، جدول ماهانه، دورهٔ ناقص، تحلیل پوشه و
اجرای مدل در پروفایل تمیز را پوشش می‌دهد؛ `tests/test_gui_bridge.py` انتقال درخواست
فارسی به workflow را بررسی می‌کند؛ `tests/test_modern_ui.py` پنجرهٔ واقعی، Worker،
خروجی Excel، فونت/تم، گفت‌وگو، پنل سخت‌افزار و ویجت شناور را بدون Mock آزمون می‌کند.

پوشش: نرمال‌سازی (شامل واحد مقیاس)، مترادف مؤسسات، خواندن/بازرسی، ساختار و جدول،
تشخیص معنایی ترکیبی، محاسبات، **آمار توصیفی/فراوانی/گروه/متقاطع/روند/همبستگی/پرت/کیفیت**،
تطبیق اختیاری وام، ناهنجاری، لایهٔ AI (با سرور HTTP واقعی برای Ollama و ComfyUI) و سیاست
مقایسهٔ دقیق Decimal، پایش، نویسندهٔ گزارش، جریان کامل نمونهٔ V4.2، فایل‌های عمومی، CLI،
پل Thread/GUI، چیدمان بسته، و موارد مرزی.

تست‌ها هیچ Mock ای برای منطق داخلی ندارند؛ فقط دو سرویس بیرونی (Ollama/ComfyUI) با سرور
HTTP محلی و واقعی شبیه‌سازی شده‌اند.

---

## ۱۰. مقایسه با نمونهٔ V4.2 و گزارش مرجع V5

| شاخص | مقدار V4.2 | مقدار موتور جدید | مقدار گزارش مرجع V5 |
|---|---|---|---|
| تعداد وام | ۶ | ۶ | ۶ |
| مجموع اصل وام | 325,000,000 | 325,000,000 | 325,000,000 |
| مجموع اقساط ماهانه | 14,749,000 | 14,749,000 | 14,749,000 |
| مجموع تعهدات ماهانه | 18,999,000 | 18,999,000 | 18,999,000 |
| مجموع بازپرداخت قراردادی | 347,576,000 | 347,576,000 | 347,576,000 |
| تعهد خارج از جدول وام | 4,250,000 | 4,250,000 | 4,250,000 |

این جدول فقط **ماژول اختیاری وام** را روی نمونهٔ Book2 نشان می‌دهد. مسیر عمومی محصول به
این اعداد وابسته نیست و برای هر جدول اداری، آماری و سازمانی کار می‌کند.

اعداد V5 از فایل `reference/v5_sample_report/` خوانده شده‌اند و با اجرای واقعی این بسته
روی همان ورودی یکسان‌اند. این اعداد از فایل استخراج و محاسبه می‌شوند؛ هیچ‌کدام در کد ثابت
نشده‌اند. خروجی Book2 در این نسخه ۲۲ برگه دارد: برگه‌های عمومی آماری به‌همراه برگه‌های
اختیاری وام/تعهد. برای فایل عمومی بدون وام، برگه‌های حسابرسی و تطبیق وام ساخته نمی‌شوند.

---

## ۱۱. محدودیت‌های شناخته‌شده

1. GPU/VRAM فقط در ویندوز و با شمارنده‌های `typeperf` قابل اندازه‌گیری است؛ در این محیط
   (Linux) مسیر «در دسترس نبود» تست شده و مسیر واقعی تأیید نشده است.
2. انتساب VRAM به یک پردازش تنها وقتی ممکن است که نام نمونهٔ شمارنده شامل PID باشد؛
   درایورهای دیگر این اطلاعات را نمی‌دهند و همین در گزارش نوشته می‌شود.
3. فایل‌های `.xls` قدیمی و فایل‌های دارای ماکرو فقط به‌صورت خواندنی پشتیبانی می‌شوند؛
   ماکرو هرگز اجرا نمی‌شود و برای `.xls` پیام صریح داده می‌شود.
4. شناسایی ساختار ابتکاری و امتیازمحور است؛ برای چیدمان‌های بسیار نامنظم، اطمینان پایین
   گزارش و تأیید کاربر گرفته می‌شود (به‌جای حدس قطعی).
5. تطبیق نام‌ها مبتنی بر یکسان‌سازی و شباهت است؛ «تطبیق احتمالی» هرگز به‌عنوان قطعی ثبت
   نمی‌شود و نیازمند تأیید انسانی است.
6. رابط گرافیکی و لانچر ویندوزی در این محیط اجرا نشده‌اند (نبود Tk و ویندوز). قرارداد
   عدم دسترسی Thread به Widgetها با `UiBridge`، `Queue`، `root.after` و تست AST بررسی شده
   است، اما رندر واقعی پنجره باید روی ویندوز تأیید شود.
7. تفسیر «درصد GPU» عمداً به‌صورت «بیشترین درصد یک Engine» گزارش می‌شود، نه درصد کل GPU.
8. سلولی که فقط متن فرمول دارد و مقدار محاسبه‌شدهٔ ذخیره‌شده ندارد (فایل ذخیره‌نشده در Excel)
   عدد تلقی نمی‌شود؛ به‌جای حدس، به‌عنوان متن با هشدار ثبت می‌شود.
9. تحلیل آماری اکتشافی است و آزمون استنباطی (t، ANOVA، کای‌دو، رگرسیون) را خودکار اجرا
   نمی‌کند؛ برای نتیجه‌گیری علّی یا آزمون فرض، طرح نمونه‌گیری و آزمون مناسب لازم است.
10. همبستگی و روند توصیفی‌اند و علت‌ومعلول را اثبات نمی‌کنند. گروه‌های کوچک نیز در
    گزارش به‌عنوان محدودیت ذکر می‌شوند.
11. انتخاب خودکار تحلیل بر اساس نوع داده و اندازهٔ دسته‌هاست؛ کاربر می‌تواند با `--task`
    و `--goal` جهت تحلیل را مشخص کند.

---

## ۱۲. نقشهٔ مهاجرت از V4.2

| در V4.2 | در بستهٔ جدید |
|---|---|
| خواندن Workbook و پیمایش برگه‌ها | `excel/reader.py`, `excel/workbook_inspector.py` |
| تشخیص سرستون/جدول | `excel/sheet_analyzer.py`, `excel/table_detector.py` |
| استخراج وام/تعهد و جمع‌های اعلامی | `excel/data_extractor.py` |
| تبدیل اعداد و تاریخ فارسی | `excel/normalization.py` |
| جمع/میانگین/بازپرداخت | `excel/calculations.py` |
| تطبیق و مغایرت | `excel/reconciliation.py` |
| هشدارهای داده | `excel/anomaly_detection.py` |
| داشبورد و برگه‌های گزارش | `excel/report_writer.py` |
| کلاینت Ollama و انتخاب مدل | `ai/ollama_client.py` |
| ساخت Prompt و خلاصهٔ فارسی | `ai/prompt_manager.py`, `ai/context_builder.py` |
| کنترل خروجی مدل | `ai/response_validator.py` |
| پایش CPU/RAM/GPU و گزارش عملکرد | `monitoring/*` |
| استودیو تصویر ComfyUI | `ai/image_studio.py` + پنجرهٔ «استودیو تصویر» در GUI |
| رابط Tkinter و لانچر BAT | `ui/wizard.py`, `Run_Local_Office_AI*.bat` |

---

## ۱۳. گام‌های بعدی پیشنهادی

1. اجرای واقعی روی ویندوز برای تأیید مسیر GPU/VRAM و رابط گرافیکی.
2. افزودن پروفایل‌های اختصاصی برای صورت‌حساب بانکی و فیش حقوقی (فقط با تنظیمات، بدون
   تغییر موتور).
3. امضای دیجیتال خروجی و بایگانی نسخه‌دار گزارش‌ها.
4. اتصال به سایر ابزارهای دستیار اداری از طریق لایهٔ `app/coordinator.py`.

---

## ۱۴. فایل‌های قدیمی و مرجع

- `legacy/Ollama_Excel_Assistant_V4_2.py` و `legacy/Run_Ollama_Excel_Assistant_V4_2.bat`
  دست‌نخورده نگه داشته شده‌اند، اما لانچر اصلی نیستند و از `py` سراسری استفاده می‌کنند.
- `data/input/Book2_AI_V4_2_Monitored.xlsx` نمونهٔ نمایشی ماژول اختیاری وام است.
- `data/input/General_Office_Demo.xlsx` نمونهٔ دادهٔ اداری/کارکنان برای مسیر عمومی است.
- `reference/v5_sample_report/` گزارش اجرایی نمونهٔ V5 را فقط به‌عنوان مرجع طراحی نگه
  می‌دارد و در زمان اجرا خوانده نمی‌شود.

---

## ۱۵. تغییرات این نسخهٔ ادغامی

| فایل | تغییر |
|---|---|
| `excel/normalization.py` | واحدهای مقیاس هزار/میلیون/میلیارد/همت، مترادف مؤسسات، `extract_numeric_claims`، بارگذاری `synonyms.json` |
| `excel/sheet_analyzer.py` | تشخیص نقش با عنوان + محتوا، گزارش تناقض عنوان/داده، برچسب فارسی نقش‌ها |
| `excel/table_detector.py` | انتقال دلایل نقش‌ها به `TableCandidate.reasons` |
| `excel/data_extractor.py` | اعمال مترادف‌ها روی `institution_key` بدون تغییر `institution_raw` |
| `excel/reconciliation.py` | استفاده از مترادف‌ها در تطبیق احتمالی |
| `ai/response_validator.py` | `NumberComparisonPolicy` با اختلاف صفر، مقایسهٔ Decimal، پشتیبانی واحد مقیاس |
| `ai/context_builder.py` | افزودن شمارش تطبیق و فهرست مؤسسات به facts |
| `ai/prompt_manager.py` | خلاصهٔ اجرایی قطعی هم‌سطح گزارش V5 |
| `excel/statistical_analysis.py` | موتور عمومی آمار توصیفی، فراوانی، گروه، متقاطع، روند، همبستگی، پرت و کیفیت |
| `excel/generic_report.py` | برگه‌های گزارش عمومی (شناسنامه، توصیفی، فراوانی، گروه، روند/همبستگی، پرت، روش‌شناسی) |
| `excel/report_writer.py` | ساخت شرطی برگه‌ها، داشبورد عمومی، خلاصه اجرایی عمومی و سیاست مقایسه |
| `app/workflow.py`, `core/models.py` | تحلیل عمومی روی جدول اصلی، ماژول اختیاری وام، `AnalysisResult.generic` |
| `ai/context_builder.py`, `ai/prompt_manager.py` | بستهٔ واقعیت عمومی، اعداد مجاز آماری، Prompt و خلاصهٔ قطعی عمومی |
| `ui/event_bridge.py` | پل Queue/Event بین Thread و Tk |
| `ui/wizard.py` | حذف دسترسی مستقیم Thread به Widgetها؛ `root.after` تنها نقطهٔ به‌روزرسانی |
| `core/settings.py` | `synonyms_file` و `synonyms` با مسیر مشتق‌شده از ریشه |
| `config/synonyms.json` | فایل مترادف‌های قابل ویرایش |
| `Run_Local_Office_AI.bat`, `Run_Local_Office_AI_GUI.bat`, `Run_Tests.bat` | فقط `venv\Scripts\python.exe`؛ بدون `py` سراسری |
| `legacy/`, `data/input/`, `reference/v5_sample_report/` | جداسازی فایل‌های قدیمی، نمونه و مرجع |
| `tests/test_units_synonyms_semantics.py`, `tests/test_ai_exact_policy.py`, `tests/test_gui_bridge.py`, `tests/test_packaging.py` | تست‌های متمرکز قابلیت‌های جدید |
