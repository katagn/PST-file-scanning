📧 PST Email Extractor (.eml Exporter)

Šī ir Windows WPF lietotne, kas analizē Microsoft Outlook PST failus un eksportē tās e-pasta vēstules, kurām ir liels skaits saņēmēju (pēc noklusējuma ≥ 20). Eksportētie e-pasti tiek saglabāti .eml formātā lietotāja izvēlētā mapē.
Lietotnes izveidei tika lietota Independentsoft.PST bibliotēkas testa versija. Lietotne tika testēta ar mākslīgi ģenerētiem PST failiem, tika konstatēts, ka neņems pretī lielākus failus kā 2,5 GB. Nav testēta ar reāliem failiem vēl.
🚀 Funkcionalitāte

    ✅ Atbalsta vairāku PST failu izvēli un apstrādi

    ✅ Saskaita visus e-pastus visos izvēlētajos PST failos

    ✅ Apstrādā katru e-pastu un eksportē tos, kuros ir 20 vai vairāk adresātu (To + Cc)

    ✅ Saglabā e-pastus .eml formātā uz diska

    ✅ Rāda detalizētu progresu, ieskaitot:

        Apstrādāto e-pastu skaitu

        Kopējo e-pastu skaitu

        Aptuveno atlikušo laiku

        Aktīvo failu

    ✅ Apstrādā arī dziļi ligzdotas mapes

    ✅ Nodrošina kļūdu apstrādi un izvadīšanu uz konsoles

🖼 Lietotāja interfeiss

    Pogas PST failu un izvadmapes izvēlei

    Progresu josla apstrādes laikā

    Teksts ar statusa informāciju

    Paziņojumi par kļūdām vai statusa ziņojumi (MessageBox)

🧾 Tehniskās detaļas

    Uzbūvēts ar C# WPF

    Izmanto Independentsoft.PST bibliotēku PST failu lasīšanai

    Multivītņu apstrāde (Task.Run, Dispatcher.Invoke, Interlocked.Increment)

    Drošs failu nosaukumu ģenerēšanas mehānisms (DateTime.Now.Ticks + Subject)

    Atbalsta angļu un daļēji latviešu valodu UI ziņojumos

⚙️ Sistēmas prasības

    Windows OS

    .NET 6/7/8 (atkarībā no projekta konfigurācijas)

    PST failu paraugi testēšanai (nav iekļauti projektā)

    Licencēta vai izmēģinājuma versija Independentsoft.PST bibliotēkai

📦 Kā lietot

    Palaidiet lietotni

    Spiediet "Choose PST files" pogu un izvēlieties vienu vai vairākus .pst failus

    Spiediet "Choose folder" un izvēlieties mapi, kur saglabāt .eml failus

    Spiediet "Start processing"

    Pēc apstrādes beigām saņemsiet paziņojumu un .eml faili būs saglabāti norādītajā mapē
