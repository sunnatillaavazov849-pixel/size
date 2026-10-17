# Size — CorelDRAW uchun o'lchamlarni avtomatik ko'rsatuvchi plagin

Obyekt yoki harflarni tanlaysiz, `Size` makrosini ishga tushirasiz va plagin
ish sahifasining o'zida quyidagilarni chizib beradi:

- har bir tanlangan obyekt yoki harfning **eni** va **bo'yi**;
- **har bir qo'shni harf orasidagi bo'shliq**, so'zlar orasidagi bo'shliq ham.
  Harflar bitta matn obyekti ichida bo'lsa yoki yozuv egri chiziqqa
  aylantirilib bitta obyekt qilib birlashtirilgan bo'lsa ham, plagin harflarni
  alohida aniqlaydi;
- oraliq markazlardan emas, birinchi harfning **eng o'ngga chiqib turgan
  nuqtasidan** keyingi harfning **eng chapga chiqib turgan nuqtasigacha**
  o'lchanadi. Bunda harfning haqiqiy ko'rinadigan konturi olinadi, shriftning
  ichki bo'sh maydonlari (sidebearing) hisobga kirmaydi;
- butun yozuvning **umumiy eni va bo'yi**;
- barcha qiymatlar **santimetrda**, nuqtadan keyin **bitta raqam** bilan:
  `10.6 cm`, `25.0 cm`, `2.5 cm` (Windows'da vergul sozlangan bo'lsa ham nuqta
  ishlatiladi);
- o'lchamlar **yordamchi chiziqlar** va uchlaridagi **strelkalar** bilan
  chiziladi. Yordamchi chiziqlar aynan o'lchanayotgan harfning eng chetki
  nuqtasidan boshlanadi;
- raqamlar obyektlar va boshqa o'lcham yozuvlari bilan **ustma-ust tushmaydi**:
  joy band bo'lsa, plagin yozuvni avtomatik ravishda bo'sh joyga suradi;
- barcha o'lchamlar bir vaqtda, `Size - o'lchamlar` degan alohida qatlamda
  bitta guruh bo'lib paydo bo'ladi.

Misol: ikkita harf tanlansa, har bir harfning eni (`10.6 cm`), bo'yi
(`20.0 cm`) va ikki harf orasidagi bo'shliq (`2.5 cm`) alohida ko'rsatiladi.
"new wor" yozuvida esa n–e, e–w, w–w (so'zlar orasidagi bo'shliq), w–o va
o–r oraliqlari alohida ko'rsatiladi.

## Talablar

- CorelDRAW X7, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024 yoki undan
  yangi versiya.
- CorelDRAW o'rnatilganda **Visual Basic for Applications (VBA)** komponenti
  ham o'rnatilgan bo'lishi kerak. Agar `Alt+F11` bosilganda VBA muharriri
  ochilmasa, CorelDRAW o'rnatuvchisini qayta ishga tushirib, *Modify* orqali
  VBA'ni qo'shing.

## O'rnatish

1. CorelDRAW'ni oching va `Alt+F11` tugmalarini bosing (yoki
   *Tools → Macros → Visual Basic Editor*; 2020+ versiyalarda
   *Tools → Scripts → Script Editor*).
2. Chap tomondagi **Project** oynasida `GlobalMacros` loyihasini tanlang.
3. *File → Import File...* menyusini bosib, `src/SizePlugin.bas` faylini
   tanlang. Loyihada `SizePlugin` moduli paydo bo'ladi.
4. `Ctrl+S` bilan saqlang va VBA muharririni yoping.

## Ishlatish

1. Bir yoki bir nechta obyekt, harf yoki matnni tanlang.
2. `Alt+Shift+F11` (*Tools → Macros → Run Macro*, 2020+ versiyalarda
   *Tools → Scripts → Run Script*) oynasini oching va
   `GlobalMacros → SizePlugin.Size` makrosini ishga tushiring.

Qulay bo'lishi uchun makrosni tugma yoki klaviatura kombinatsiyasiga
biriktiring: *Tools → Options → Customization → Commands* bo'limida
ro'yxatdan **Macros** ni tanlang, `GlobalMacros.SizePlugin.Size` ni asboblar
paneliga sudrab qo'ying yoki *Shortcut Keys* yorlig'ida unga tugma
biriktiring.

### Makroslar

| Makros         | Vazifasi |
|----------------|----------|
| `Size`         | Tanlangan obyektlarni o'lchaydi va o'lchamlarni chizadi. |
| `SizeAsk`      | Sozlamalar oynasini ochadi. **Saqlash va o'lchash** bosilsa, tanlangan obyektlarni darhol o'lchaydi. |
| `SizeSettings` | Sozlamalar oynasi (pastga qarang). |
| `SizeClear`    | Joriy sahifadagi barcha Size o'lchamlarini (qatlami bilan birga) o'chiradi. |

Har bir `Size` ishga tushishi bitta qadam sifatida yoziladi, shuning uchun
`Ctrl+Z` uni to'liq bekor qiladi. Natija alohida qatlamda bo'lgani uchun uni
*Object Manager* orqali yashirish, chop etishdan chiqarish yoki qulflash
mumkin. Qayta ishga tushirilganda yangi o'lchamlar oldingi o'lchamlar
ustiga tushmaydi.

## Sozlamalar oynasi (`SizeSettings`)

Oyna to'q kulrang fonda, yozuvlari oq. Unda quyidagilar bor:

| Bo'lim | Sozlama | Standart qiymat | Izoh |
|--------|---------|-----------------|------|
| Raqamlar | Shrift | `Arial` | Kompyuterga o'rnatilgan shriftlar ro'yxatidan tanlanadi. Ro'yxatda yo'q shrift yozilsa, Arial ishlatiladi. |
| | Qalin (bold) | yoqilgan | |
| | Shrift o'lchami (pt) | Avto | **Avto** belgisini olib tashlang va o'lchamni yozing, masalan `48`. |
| Chiziqlar | Chiziq qalinligi (mm) | Avto | O'lcham chizig'ining qalinligi. Yordamchi chiziqlar undan biroz ingichka (2/3) bo'ladi. |
| | Strelka uzunligi (mm) | Avto | Strelka eni uzunligiga mos ravishda o'zgaradi. |
| | Rang (R,G,B) | `0,90,200` | Chiziqlar, strelkalar va raqamlar rangi. **Tanlash** tugmasi CorelDRAW'ning rang tanlash oynasini ochadi. |
| O'lchash | Harflarni alohida aniqlash | yoqilgan | Matn, egri chiziqqa aylantirilgan yozuv yoki guruh harflarga ajratiladi. O'chirilsa, har bir obyekt butunligicha o'lchanadi. |
| | Har bir harfning eni va bo'yi | yoqilgan | O'chirilsa, faqat oraliqlar va umumiy o'lcham ko'rsatiladi. Uzun yozuvlarda chizma ancha ixcham bo'ladi. |
| | Umumiy eni va bo'yi | yoqilgan | Butun tanlovning umumiy o'lchami. |
| | Kontur qalinligini hisobga olish | o'chirilgan | Obyekt konturining (outline) qalinligi o'lchamga qo'shiladi. |
| | Bo'yi raqamini vertikal yozish | yoqilgan | O'chirilsa, bo'yi raqamlari gorizontal yoziladi. |

Tugmalar: **Saqlash va o'lchash**, **Saqlash**, **Bekor qilish**. Sozlamalar
Windows reyestrida saqlanadi va CorelDRAW qayta ochilganda ham saqlanib
qoladi.

Oyna birinchi marta ochilganda plagin uni `GlobalMacros` loyihasida
`SizeSettingsForm` nomi bilan o'zi yaratadi. Uni o'chirmang va
o'zgartirmang. CorelDRAW yopilayotganda `GlobalMacros` ni saqlashni so'rasa,
**Ha** ni bosing. Agar oynani yaratib bo'lmasa (masalan, VBA loyihasi
qulflangan bo'lsa), sozlamalar avvalgidek ketma-ket so'rov oynalari orqali
so'raladi.

### Avtomatik shrift o'lchami

Shrift o'lchami `0` bo'lsa, u tanlangan harflarning balandligiga qarab
tanlanadi:

- kichik matnda raqamlar harflardan **yirikroq** bo'ladi (taxminan 1.5
  barobar, eng ko'pi 28 pt);
- katta harflarda (masalan, peshtaxta va reklama yozuvlari) raqamlar harf
  balandligining ~13 % iga teng bo'ladi. Ular uzoqdan ham aniq ko'rinadi,
  lekin chizmani bosib ketmaydi.

Chiziq qalinligi va strelka o'lchami **Avto** bo'lsa, ular ham masofalar
kabi shrift o'lchamiga mos ravishda kattalashadi yoki kichrayadi.

## Qanday ishlaydi

- **O'lchash.** Har bir obyektning ko'rinadigan konturi o'qiladi. Matn uchun
  ham shunday, uni egri chiziqqa aylantirish shart emas. Har bir kontur
  alohida o'lchanadi, so'ng konturlar harflarga yig'iladi: "o" ichidagi
  teshik, "i" va "j" nuqtasi, "Ö" ustidagi belgilar harfning o'ziga
  qo'shiladi. Asl obyektga tegilmaydi.
- **Oraliqlar.** Har bir harf uchun uning eng o'ngga, eng chapga, eng tepaga
  va eng pastga chiqib turgan nuqtalari aniqlanadi. Qo'shni harflar orasidagi
  bo'shliq shu nuqtalar orasidagi gorizontal (ustma-ust turganda vertikal)
  masofa sifatida o'lchanadi. Masalan "j" uchun eng chap nuqta pastki
  ilmog'ining uchi bo'ladi. Agar ikki harf orasida uchinchisi tursa, oraliq
  har bir qo'shni juftlik uchun alohida ko'rsatiladi.
- **Yordamchi chiziqlar** aynan shu chetki nuqtalardan boshlanadi. Iloji
  bo'lsa, oraliq strelkalari ikkala harfning chetki nuqtasiga to'g'ridan-to'g'ri
  tegib turadi.
- **Joylashtirish.** Eni odatda obyektning tepasiga, bo'yi bo'sh joyi ko'p
  tomonga, oraliq esa iloji bo'lsa ikki obyekt o'rtasiga, bo'lmasa pastga
  qo'yiladi. Har bir o'lcham uchun ko'plab variantlar tekshiriladi va raqami
  obyektlar hamda boshqa raqamlar bilan kesishmaydigan eng yaqin variant
  tanlanadi.

## Cheklovlar

- Egri segmentlarning eng chetki nuqtalari tugun (node) bilan belgilanmagan
  bo'lsa (sifatsiz shriftlarda uchraydi), yordamchi chiziq harfning butun
  balandligi bo'ylab qo'yiladi.
- Bitta egri chiziqdan iborat logotip ham harflarga ajratish rejimida
  konturlar bo'yicha bo'linadi. Buni xohlamasangiz, `SizeSettings` orqali
  harflarga ajratishni o'chiring.
- Hujjatda chizma masshtabi (*Drawing scale*) sozlangan bo'lsa ham,
  o'lchamlar haqiqiy (1:1) santimetrda ko'rsatiladi.
- Juda zich joylashgan kichik harflarda raqamlar uchun joy kam bo'ladi.
  Bunday holda raqamlar baribir ustma-ust tushmaydi, lekin ba'zi yordamchi
  chiziqlar uzunroq bo'lib, boshqa obyektlar ustidan o'tishi mumkin. Shrift
  o'lchamini `SizeAsk` orqali kichraytirsangiz, chizma ixchamroq bo'ladi.
