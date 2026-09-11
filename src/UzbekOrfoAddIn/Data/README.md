# Ilova bilan birga keladigan maʼlumotlar

Bu fayllar qoʻshimcha bilan tarqatiladigan, versiyalangan boshlangʻich maʼlumotlardir.

- `uzbek_dictionary.json` — asosiy manba (`uzbekorfo-dictionary-v1`). `Entries` va metadataʼni shu yerda tahrirlang; bu build paytidagi manba boʻlib, yakuniy foydalanuvchiga yuborilmaydi.
- `uzbek_main.dic` — UTF-8 kodlashdagi, har satrda bitta soʻzdan iborat hosil qilinadigan imlo lugʻati. Uni qoʻlda tahrirlamang.
- `uzbek_dictionary_metadata.json` — faqat taʼrif, qoida, grammatik izoh yoki misolga ega yozuvlardan tuzilgan ixcham, hosil qilinadigan runtime indeks. Uni qoʻlda tahrirlamang.
- `grammar_rules.json`, `uzbek_suffixes.json` va `proper_nouns.json` grammatika tahlilini qoʻllab-quvvatlaydi.
- `translit_exceptions.json` va `explanations.json` foydalanuvchi kompyuteridagi tegishli saqlagichlar uchun boshlangʻich maʼlumot beradi.
- `uzbek_freq_seed.json` — til modeli ustida ishlash uchun boshlangʻich manba.

Loyiha `UzbekOrfoAddIn.csproj` faylida koʻrsatilgan runtime maʼlumotlarini ilovaga joylaydi va chiqish papkasiga nusxalaydi. Maʼlumotlar toʻplamini almashtirganda, barqaror fayl nomini saqlang va uning kelib chiqishi hamda sanasini shu hujjatda yoki pull requestʼda qayd eting.

Lugʻat eksport formatlari: JSON — toʻliq koʻchirish va zaxira nusxa formati. DIC — UTF-8 kodlashdagi, har satrda bitta soʻzdan iborat imlo lugʻati; unda taʼriflar va boshqa metadata ataylab saqlanmaydi. Excel formatlari tuzilgan maydonlarni saqlab qoladi.

Build jarayoni paketlashdan oldin ikkala runtime faylini ham hosil qiladi. Imlo tekshiruvi faqat DICʼni yuklaydi; taʼrif qidiruvi mahalliy izoh topilmagandagina ixcham metadata indeksini yuklaydi. Ilova bilan keladigan fayllar paket ichidagi baytlari oʻzgarganda avvaldan oʻrnatilgan nusxada yangilanadi, biroq foydalanuvchi lugʻatlari va foydalanuvchi izohlari hech qachon ustiga yozilmaydi. Katta importdan keyin `eng\format-dictionary-source.ps1`, manbani tahrirlagandan keyin `eng\sync-dictionary-data.ps1`, ilovaga biriktirilgan maʼlumotlarni oʻzgartirishdan yoki CIʼda tekshirishdan oldin esa `eng\validate-dictionary-sync.ps1` ni ishga tushiring. Generator takroriy, boʻsh, koʻrinmas yoki boshqaruv belgilariga ega soʻzlarni rad etadi.
