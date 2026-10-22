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

Plagin uchta fayldan iborat (hammasi `src/` papkasida):

| Fayl | Nima |
|------|------|
| `SizePlugin.bas` | Plaginning o'zi (o'lchash, joylashtirish, sozlamalar). |
| `SizeForm.frm` | Size oynasi. |
| `SizeForm.frx` | Oynaning ikkilik qismi. `SizeForm.frm` bilan **bir papkada** turishi shart. |

1. Fayllarni GitHub'dan **Code → Download ZIP** orqali yuklab oling va
   arxivni oching. `.frm`/`.bas` fayllarni brauzerdan nusxa ko'chirib yoki
   "raw" ko'rinishda saqlamang, aks holda qator oxirlari buziladi va VBA ularni
   to'g'ri import qilmaydi.
2. CorelDRAW'ni oching va `Alt+F11` tugmalarini bosing (yoki
   *Tools → Macros → Visual Basic Editor*; 2020+ versiyalarda
   *Tools → Scripts → Script Editor*).
3. Chap tomondagi **Project** oynasida `GlobalMacros` loyihasini tanlang.
4. *File → Import File...* orqali avval `src/SizePlugin.bas`, keyin
   `src/SizeForm.frm` faylini import qiling. Loyihada `Modules → SizePlugin`
   va `Forms → SizeForm` paydo bo'ladi.
5. `Ctrl+S` bilan saqlang va VBA muharririni yoping.

Eski versiyadan yangilashda avval eski `SizePlugin` modulini va (bo'lsa)
`SizeForm` oynasini o'chiring (o'ng tugma → *Remove*), keyin yangilarini
import qiling. Oldingi versiya avtomatik yaratgan `SizeSettingsForm` endi
kerak emas, uni ham o'chirsa bo'ladi.

## Ishlatish

Plaginda bitta makros bor: **`Size`**. U Size oynasini ochadi, hammasi shu
oynada:

1. `Alt+Shift+F11` (*Tools → Macros → Run Macro*, 2020+ versiyalarda
   *Tools → Scripts → Run Script*) oynasini oching va
   `GlobalMacros → SizePlugin.Size` ni ishga tushiring. Size oynasi ochiladi.
2. Chizmada obyekt, harf yoki matnni tanlang. Oyna ochiq turaveradi.
3. **O'lchash** tugmasini (yoki `Enter`) bosing. Natija oynaning pastki
   qatorida chiqadi, masalan: `Tayyor: 19 ta o'lcham`.
4. Boshqa obyektlarni tanlab, yana **O'lchash** ni bosish mumkin.
   **Tozalash** joriy sahifadagi barcha Size o'lchamlarini o'chiradi,
   **Yopish** oynani yopadi.

Qulay bo'lishi uchun `Size` ni tugma yoki klaviatura kombinatsiyasiga
biriktiring: *Tools → Options → Customization → Commands* bo'limida
ro'yxatdan **Macros** ni tanlang, `GlobalMacros.SizePlugin.Size` ni asboblar
paneliga sudrab qo'ying yoki *Shortcut Keys* yorlig'ida unga tugma
biriktiring.

Har bir **O'lchash** bitta qadam sifatida yoziladi, shuning uchun
`Ctrl+Z` uni to'liq bekor qiladi. Natija alohida qatlamda bo'lgani uchun uni
*Object Manager* orqali yashirish, chop etishdan chiqarish yoki qulflash
mumkin. Qayta ishga tushirilganda yangi o'lchamlar oldingi o'lchamlar
ustiga tushmaydi.

## Size oynasi

Oyna minimalistik: och kulrang fon, oq kartalar, yumaloq burchaklar. Har bir
maydon ustiga sichqoncha olib borilsa, qisqa izoh chiqadi.

| Karta | Maydon | Standart | Vazifasi |
|-------|--------|----------|----------|
| Raqamlar | Shrift | `Arial` | Ro'yxatdan tanlang yoki istalgan shrift nomini yozing. Kompyuterda bo'lmasa, CorelDRAW uni o'zi almashtiradi. |
| | Hajm, pt | Avto | `−` / `+` yoki qiymatni yozing. `−`/`+` bosilsa Avto o'chadi. |
| | Qalin | yoqilgan | Qalin raqamlar. |
| Chiziqlar | Qalinlik, mm | Avto | O'lcham chizig'i. Yordamchi chiziqlar 2/3 qalinlikda. |
| | Strelka, mm | Avto | Strelka uzunligi. |
| | Rang | `0,90,200` | R,G,B. **Tanlash** — CorelDRAW rang oynasi. |
| Ko'rsatish | Harflarga ajratish | yoqilgan | Matn va egri chiziqdagi har bir harf alohida. |
| | Eni va bo'yi | yoqilgan | Har bir harfning eni va bo'yi. |
| | Umumiy o'lcham | yoqilgan | Butun tanlovning eni va bo'yi. |
| | Kontur qalinligi | o'chirilgan | Kontur qalinligi o'lchamga qo'shiladi. |
| | Vertikal raqam | yoqilgan | Bo'yi raqami 90° burilib yoziladi. |

Tugmalar: **O'lchash** (yoki `Enter`), **Tozalash**, **Yopish**. Natija yoki
xato tugmalar ostidagi qatorda chiqadi (`Tayyor: 19 ta o'lcham`,
`Obyekt tanlanmagan.`). Sozlamalar **O'lchash** yoki **Yopish** bosilganda
Windows reyestriga saqlanadi.

Kartalar, maydonlar, tugmalar va almashtirgichlarning yumaloq fonlari kichik
rasm sifatida chiziladi va `%TEMP%\SizeUI` papkasida saqlanadi (faqat birinchi
marta). Bu papkaga yozib bo'lmasa (masalan, Windows foydalanuvchi nomi
kirillcha bo'lsa), `%LOCALAPPDATA%\SizeUI` yoki `C:\SizeUI` ishlatiladi; hech
biri bo'lmasa, oyna yumaloq burchaksiz, oddiy ranglarda ochiladi.
**Yopish** oynani yopmaydi, yashiradi, shuning uchun keyingi ochilishlar darhol
bo'ladi.

Shrift maydoni CorelDRAW'ning shriftlar ro'yxatini o'qimaydi: ko'p shrift
o'rnatilgan kompyuterlarda bu CorelDRAW'ni qotirib qo'yardi.

Agar oyna baribir qotsa, shu papkadagi `size.log` faylini yuboring: unda har
bir qadam vaqti bilan yoziladi va qayerda to'xtagani ko'rinadi.

Agar CorelDRAW ochilgandan keyin *birinchi* makros har doim sekin ishga
tushsa, bu CorelDRAW'ning VBA'ni kechiktirib yuklash sozlamasidan:
*Tools → Options → Global → VBA → Delay load VBA*. Belgini olib tashlasangiz,
makroslar darhol ishga tushadi, lekin CorelDRAW o'zi biroz sekinroq ochiladi.

Agar `SizeForm` import qilinmagan bo'lsa, plagin oynani VBA muharriri orqali
o'zi yaratishga urinadi (`SizeSettingsForm`). Bu ham bo'lmasa, `Size` oddiy
so'rov oynasini ochadi: `1` — o'lchash, `2` — sozlamalar, `3` — o'lchamlarni
o'chirish.

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
  konturlar bo'yicha bo'linadi. Buni xohlamasangiz, Size oynasida
  harflarga ajratishni o'chiring.
- Hujjatda chizma masshtabi (*Drawing scale*) sozlangan bo'lsa ham,
  o'lchamlar haqiqiy (1:1) santimetrda ko'rsatiladi.
- Juda zich joylashgan kichik harflarda raqamlar uchun joy kam bo'ladi.
  Bunday holda raqamlar baribir ustma-ust tushmaydi, lekin ba'zi yordamchi
  chiziqlar uzunroq bo'lib, boshqa obyektlar ustidan o'tishi mumkin. Shrift
  o'lchamini Size oynasida kichraytirsangiz, chizma ixchamroq bo'ladi.
