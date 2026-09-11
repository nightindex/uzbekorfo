# Uzbek Orfo qoʻshimchasi

Uzbek Orfo — Microsoft Word uchun oʻzbekcha imlo va grammatika tekshiruvi, lugʻat boshqaruvi hamda yozuvlar orasida oʻgirish imkoniyatini beruvchi VSTO qoʻshimchasidir.

## Imkoniyatlari

- Xatolar roʻyxati bilan imlo va grammatika tekshiruvi
- Takliflar va barcha mos holatlarni almashtirish
- Shaxsiy lugʻatni boshqarish: soʻz qoʻshish, import va eksport qilish
- Soʻz taʼrifi va izohini qidirish
- Lotin va kirill yozuvlari orasida oʻgirish yordamchilari
- Xatolar va hisobotlarni eksport qilish

## Talablar

- Windows
- Microsoft Word (`Office 2016+` tavsiya etiladi)
- Office/VSTO vositalari oʻrnatilgan Visual Studio
- .NET Framework 4.7.2

## Yigʻish va ishga tushirish

1. Visual Studioʼda `UzbekOrfoAddIn.slnx` faylini oching.
2. `Directory.Build.props.example` faylidan Git kuzatmaydigan `Directory.Build.props` nusxasini yarating va mahalliy imzolash sertifikatingizni sozlang. VSTO ishga tushadigan qoʻshimcha uchun, jumladan mahalliy nosozliklarni tuzatish jarayonida ham, imzolangan manifestlarni talab qiladi.
3. Loyihani `Debug` yoki `Release` rejimida yigʻing.
4. Wordʼni qoʻshimcha bilan ishga tushirish uchun debuggingʼni boshlang.

Toʻliq yigʻish va debugging jarayoni uchun VSTO workload hamda kompyuterga mos Microsoft Word oʻrnatilgan boʻlishi kerak.

## Loyiha tuzilmasi

- `src/UzbekOrfoAddIn/` — VSTO host va ilova manba kodi
- `eng/` — repozitoriy tekshiruvlari va muhandislik skriptlari
- `tests/` — hostga bogʻliq boʻlmagan mantiq uchun avtomatlashtirilgan testlar
- `docs/` — arxitektura va hissa qoʻshuvchilar uchun hujjatlar
- `.github/workflows/` — uzluksiz integratsiya tekshiruvlari

`src/UzbekOrfoAddIn/` ichida `Core/` shartnomalarni, `Models/` domen maʼlumotlarini, `Services/` ilova mantiqini, `Forms/` va `UI/` interfeys kodini, `Data/` esa ilova bilan birga keladigan lugʻatlar va qoidalarni saqlaydi.

## Eslatmalar

- `Data/` fayllari resurs sifatida ilovaga joylanadi va debugging uchun chiqish papkasiga nusxalanadi.
- Imzolash kalitlari (`*.pfx`) va mahalliy `Directory.Build.props` fayli repozitoriy `.gitignore` qoidalari orqali Gitʼga kiritilmaydi. Ildizdagi `Directory.Build.props.example` faylidan mahalliy `Directory.Build.props` nusxasini yarating va release build yaratishdan oldin imzolash qiymatlarini sozlang. Sertifikat fayli yoki thumbprintʼni loyiha fayliga qoʻshmang.

## Reliz tayyorlash

Ishlab chiqarish uchun reliz — imzolangan ClickOnce `Release` buildʼidir. Imzolash sozlamalari, offline va web yangilash kanallari, kerakli release branch hamda bitta reliz-tekshiruv buyrugʻi uchun [reliz jarayoni](docs/release.md) hujjatiga qarang. Publish natijalari ataylab Git tomonidan eʼtiborga olinmaydi.

## Tekshirish

Repozitoriy ildizida quyidagi tekshiruvlarni ishga tushiring:

```powershell
.\eng\preflight.ps1
.\eng\test-ui.ps1
.\eng\test-unit.ps1
```

Test loyihasi helper mantiqini, sozlamalarni atomik yozishni hamda xatoni tuzatish va belgilash xavfsizligini xotiradagi Word double orqali tekshiradi. U Wordʼni talab qilmaydi va COM integratsiyasi testlarining oʻrnini bosmaydi. VSTO buildʼini tekshirish uchun Office development workload oʻrnatilgan Windows kompyuteri talab qilinadi.

Windows UI test vositasi Visual Studio C# kompilyatorini talab qiladi, biroq Wordʼni talab qilmaydi. U DPI masshtablanishi, scrollbar regressiyalari va kichik oynalardagi joylashuvlarni tekshiradi. Qamrovi va qoʻlda bajariladigan tekshiruvlar uchun [ekran mosligi](docs/display-compatibility.md) hujjatiga qarang.

Ilovaga biriktirilgan lugʻat fayllarining mosligini `./eng/validate-dictionary-sync.ps1` orqali tekshiring. Agar ikki hosil qilinadigan faylni yangidan yaratish kerak boʻlsa, `./eng/sync-dictionary-data.ps1` buyrugʻini ishga tushiring; u mos JSON metadataʼni saqlaydi va faqat imlo tekshiruvi uchun moʻljallangan soʻzlarga boʻsh metadata maydonlarini qoʻshadi.

Xulq-atvordagi oʻzgarishlar va Word smoke testlari haqida [hujjat xavfsizligi hamda review tuzatishlari](docs/review-fixes.md) sahifasidan oʻqing.
