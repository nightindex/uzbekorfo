namespace UzbekOrfoAddIn
{
    partial class UzbekOrfoRibbon : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public UzbekOrfoRibbon()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UzbekOrfoRibbon));
            this.tabUzbekOrfo = this.Factory.CreateRibbonTab();
            this.groupCheck = this.Factory.CreateRibbonGroup();
            this.btnCheckSpelling = this.Factory.CreateRibbonButton();
            this.btnViewErrors = this.Factory.CreateRibbonButton();
            this.btnErrorList = this.Factory.CreateRibbonButton();
            this.groupFix = this.Factory.CreateRibbonGroup();
            this.btnSuggestions = this.Factory.CreateRibbonButton();
            this.btnReplaceAll = this.Factory.CreateRibbonButton();
            this.toggleAutoCorrect = this.Factory.CreateRibbonToggleButton();
            this.groupDictionary = this.Factory.CreateRibbonGroup();
            this.btnAddToDict = this.Factory.CreateRibbonButton();
            this.btnEditDictionary = this.Factory.CreateRibbonButton();
            this.btnAddNewWords = this.Factory.CreateRibbonButton();
            this.groupTransliteration = this.Factory.CreateRibbonGroup();
            this.btnFromLatinToCyrillic = this.Factory.CreateRibbonButton();
            this.btnFromCyrillicToLatin = this.Factory.CreateRibbonButton();
            this.btnTransExceptions = this.Factory.CreateRibbonButton();
            this.groupTools = this.Factory.CreateRibbonGroup();
            this.btnCleanupSpaces = this.Factory.CreateRibbonButton();
            this.btnSwitchScript = this.Factory.CreateRibbonButton();
            this.btnSetFontTNR = this.Factory.CreateRibbonButton();
            this.groupExplanations = this.Factory.CreateRibbonGroup();
            this.btnDefinitions = this.Factory.CreateRibbonButton();
            this.btnExportErrors = this.Factory.CreateRibbonButton();
            this.groupSpecialChars = this.Factory.CreateRibbonGroup();
            this.btnSpecialChar1 = this.Factory.CreateRibbonButton();
            this.btnSpecialChar2 = this.Factory.CreateRibbonButton();
            this.btnSpecialCharAll = this.Factory.CreateRibbonButton();
            this.lblScriptIndicator = this.Factory.CreateRibbonButton();
            this.groupInfo = this.Factory.CreateRibbonGroup();
            this.btnAppInfo = this.Factory.CreateRibbonButton();
            this.tabUzbekOrfo.SuspendLayout();
            this.groupCheck.SuspendLayout();
            this.groupFix.SuspendLayout();
            this.groupDictionary.SuspendLayout();
            this.groupTransliteration.SuspendLayout();
            this.groupTools.SuspendLayout();
            this.groupExplanations.SuspendLayout();
            this.groupSpecialChars.SuspendLayout();
            this.groupInfo.SuspendLayout();
            this.SuspendLayout();
            // 
            // tabUzbekOrfo
            // 
            this.tabUzbekOrfo.Groups.Add(this.groupCheck);
            this.tabUzbekOrfo.Groups.Add(this.groupFix);
            this.tabUzbekOrfo.Groups.Add(this.groupDictionary);
            this.tabUzbekOrfo.Groups.Add(this.groupTransliteration);
            this.tabUzbekOrfo.Groups.Add(this.groupTools);
            this.tabUzbekOrfo.Groups.Add(this.groupExplanations);
            this.tabUzbekOrfo.Groups.Add(this.groupSpecialChars);
            this.tabUzbekOrfo.Groups.Add(this.groupInfo);
            this.tabUzbekOrfo.Label = "Ўзбек Орфо";
            this.tabUzbekOrfo.Name = "tabUzbekOrfo";
            // 
            // groupCheck
            // 
            this.groupCheck.Items.Add(this.btnCheckSpelling);
            this.groupCheck.Items.Add(this.btnViewErrors);
            this.groupCheck.Items.Add(this.btnErrorList);
            this.groupCheck.Label = "Текшириш";
            this.groupCheck.Name = "groupCheck";
            // 
            // btnCheckSpelling
            // 
            this.btnCheckSpelling.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnCheckSpelling.Image = ((System.Drawing.Image)(resources.GetObject("btnCheckSpelling.Image")));
            this.btnCheckSpelling.Label = "Матн текшируви";
            this.btnCheckSpelling.Name = "btnCheckSpelling";
            this.btnCheckSpelling.OfficeImageId = "SpellingAndGrammar";
            this.btnCheckSpelling.ScreenTip = "Матн текшируви  (Ctrl + Alt + Q)";
            this.btnCheckSpelling.ShowImage = true;
            this.btnCheckSpelling.SuperTip = "Ҳужжатдаги имло ва грамматик хатоларни текширади.\nИмло хатолари — қизил, граммати" +
    "к хатолар — яшил тўлқинли чизиқ билан белгиланади.\n(Грамматика текширувини ёқ/ўч" +
    " тугмаси орқали бошқариш мумкин.)";
            this.btnCheckSpelling.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCheckSpelling_Click);
            // 
            // btnViewErrors
            // 
            this.btnViewErrors.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnViewErrors.Image = ((System.Drawing.Image)(resources.GetObject("btnViewErrors.Image")));
            this.btnViewErrors.Label = "Хатоларни кўриш";
            this.btnViewErrors.Name = "btnViewErrors";
            this.btnViewErrors.OfficeImageId = "ReviewSpellingError";
            this.btnViewErrors.ScreenTip = "Хатоларни кўриш  (Ctrl + Alt + E)";
            this.btnViewErrors.ShowImage = true;
            this.btnViewErrors.SuperTip = "Охирги текширувда топилган хатолар рўйхатини кўрсатади (диалог ойнасида).\n(Хатола" +
    "рни тез кўриб чиқиш ва ўчириш учун.)";
            this.btnViewErrors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnViewErrors_Click);
            // 
            // btnErrorList
            // 
            this.btnErrorList.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnErrorList.Image = ((System.Drawing.Image)(resources.GetObject("btnErrorList.Image")));
            this.btnErrorList.Label = "Хатолар рўйхати";
            this.btnErrorList.Name = "btnErrorList";
            this.btnErrorList.OfficeImageId = "BulletedList";
            this.btnErrorList.ScreenTip = "Хатолар рўйхати  (Ctrl + Alt + W)";
            this.btnErrorList.ShowImage = true;
            this.btnErrorList.SuperTip = "Топилган барча хатоларни янги ҳужжатда батафсил рўйхат шаклида тайёрлайди (сақлаш" +
    "/чоп этиш мумкин).\n(Хатолар ҳисоботини яратиш учун.)\nТезкор тугма: Ctrl + Alt + " +
    "W.";
            this.btnErrorList.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnErrorList_Click);
            // 
            // groupFix
            // 
            this.groupFix.Items.Add(this.btnSuggestions);
            this.groupFix.Items.Add(this.btnReplaceAll);
            this.groupFix.Items.Add(this.toggleAutoCorrect);
            this.groupFix.Label = "Тузатиш";
            this.groupFix.Name = "groupFix";
            // 
            // btnSuggestions
            // 
            this.btnSuggestions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnSuggestions.Image = ((System.Drawing.Image)(resources.GetObject("btnSuggestions.Image")));
            this.btnSuggestions.Label = "Вариантлар";
            this.btnSuggestions.Name = "btnSuggestions";
            this.btnSuggestions.OfficeImageId = "AutoCorrect";
            this.btnSuggestions.ScreenTip = "Вариантлар  (Ctrl + Alt + J)";
            this.btnSuggestions.ShowImage = true;
            this.btnSuggestions.SuperTip = "Белгиланган хато сўз учун тузатиш вариантларини кўрсатади.\n(Хато сўзни тез алмашт" +
    "ириш учун.)";
            this.btnSuggestions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSuggestions_Click);
            // 
            // btnReplaceAll
            // 
            this.btnReplaceAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnReplaceAll.Image = ((System.Drawing.Image)(resources.GetObject("btnReplaceAll.Image")));
            this.btnReplaceAll.Label = "Барчасини алмаштириш";
            this.btnReplaceAll.Name = "btnReplaceAll";
            this.btnReplaceAll.OfficeImageId = "FindDialog";
            this.btnReplaceAll.ScreenTip = "Барчасини алмаштириш  (Ctrl + Alt + H)";
            this.btnReplaceAll.ShowImage = true;
            this.btnReplaceAll.SuperTip = "Танланган хато сўзнинг ҳужжатдаги барча учрашларини бирданига тўғри вариантга алм" +
    "аштиради.\n(Такрорланган хатоларни тез тузатиш учун.)";
            this.btnReplaceAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnReplaceAll_Click);
            // 
            // toggleAutoCorrect
            // 
            this.toggleAutoCorrect.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.toggleAutoCorrect.Label = "Авто тузатиш";
            this.toggleAutoCorrect.Name = "toggleAutoCorrect";
            this.toggleAutoCorrect.OfficeImageId = "AutoCorrect";
            this.toggleAutoCorrect.ScreenTip = "Авто тузатиш  (Ctrl + Alt + A)";
            this.toggleAutoCorrect.ShowImage = true;
            this.toggleAutoCorrect.SuperTip = "Автоматик тузатиш режимини ёқади/ўчиради — текширув пайтида энг яхши вариантни ўз" +
    "и қўяди.\n(Тез ва кўп ишловчи ҳужжатлар учун қулай.)";
            this.toggleAutoCorrect.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.toggleAutoCorrect_Click);
            // 
            // groupDictionary
            // 
            this.groupDictionary.Items.Add(this.btnAddToDict);
            this.groupDictionary.Items.Add(this.btnEditDictionary);
            this.groupDictionary.Items.Add(this.btnAddNewWords);
            this.groupDictionary.Label = "Луғат";
            this.groupDictionary.Name = "groupDictionary";
            // 
            // btnAddToDict
            // 
            this.btnAddToDict.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnAddToDict.Image = ((System.Drawing.Image)(resources.GetObject("btnAddToDict.Image")));
            this.btnAddToDict.Label = "Луғатга қўшиш";
            this.btnAddToDict.Name = "btnAddToDict";
            this.btnAddToDict.OfficeImageId = "DictionaryAdd";
            this.btnAddToDict.ScreenTip = "Луғатга қўшиш  (Ctrl + Alt + B)";
            this.btnAddToDict.ShowImage = true;
            this.btnAddToDict.SuperTip = "Белгиланган сўзни шахсий луғатга қўшади (кейинги текширувларда хато деб ҳисобланм" +
    "айди).\n(Тўғри, лекин кам учрайдиган сўзлар учун.)";
            this.btnAddToDict.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddToDict_Click);
            // 
            // btnEditDictionary
            // 
            this.btnEditDictionary.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnEditDictionary.Image = ((System.Drawing.Image)(resources.GetObject("btnEditDictionary.Image")));
            this.btnEditDictionary.Label = "Луғатни таҳрирлаш";
            this.btnEditDictionary.Name = "btnEditDictionary";
            this.btnEditDictionary.OfficeImageId = "Dictionary";
            this.btnEditDictionary.ScreenTip = "Луғатни таҳрирлаш  (Ctrl + Alt + G)";
            this.btnEditDictionary.ShowImage = true;
            this.btnEditDictionary.SuperTip = "Шахсий луғатдаги сўзлар рўйхатини кўриш, ўчириш ва таҳрирлаш ойнасини очиш.\n(Луға" +
    "тни бошқариш учун.)";
            this.btnEditDictionary.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnEditDictionary_Click);
            // 
            // btnAddNewWords
            // 
            this.btnAddNewWords.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnAddNewWords.Image = ((System.Drawing.Image)(resources.GetObject("btnAddNewWords.Image")));
            this.btnAddNewWords.Label = "Сўз қўшиш";
            this.btnAddNewWords.Name = "btnAddNewWords";
            this.btnAddNewWords.OfficeImageId = "FileOpen";
            this.btnAddNewWords.ScreenTip = "Сўз қўшиш  (Ctrl + Alt + Y)";
            this.btnAddNewWords.ShowImage = true;
            this.btnAddNewWords.SuperTip = "Ташқи файл (.txt, .docx ва ҳ.к.) дан янги сўзларни ўқиб, луғатга қўшади.\n(Катта л" +
    "уғат яратиш ёки ўрганиш учун.)";
            this.btnAddNewWords.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAddNewWords_Click);
            // 
            // groupTransliteration
            // 
            this.groupTransliteration.Items.Add(this.btnFromLatinToCyrillic);
            this.groupTransliteration.Items.Add(this.btnFromCyrillicToLatin);
            this.groupTransliteration.Items.Add(this.btnTransExceptions);
            this.groupTransliteration.Label = "Транслитерация";
            this.groupTransliteration.Name = "groupTransliteration";
            // 
            // btnFromLatinToCyrillic
            // 
            this.btnFromLatinToCyrillic.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnFromLatinToCyrillic.Image = ((System.Drawing.Image)(resources.GetObject("btnFromLatinToCyrillic.Image")));
            this.btnFromLatinToCyrillic.Label = "Лотиндан Кириллга";
            this.btnFromLatinToCyrillic.Name = "btnFromLatinToCyrillic";
            this.btnFromLatinToCyrillic.OfficeImageId = "Translate";
            this.btnFromLatinToCyrillic.ScreenTip = "Лотиндан Кириллга (Ctrl + Alt + Shift + C)";
            this.btnFromLatinToCyrillic.ShowImage = true;
            this.btnFromLatinToCyrillic.SuperTip = "Танланган матн ёки бутун ҳужжатни бир босишда кириллга ўтказади.\nТезкор ишлатиш: " +
    "Ctrl + Alt + Shift + C.";
            this.btnFromLatinToCyrillic.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFromLatinToCyrillic_Click);
            // 
            // btnFromCyrillicToLatin
            // 
            this.btnFromCyrillicToLatin.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnFromCyrillicToLatin.Image = ((System.Drawing.Image)(resources.GetObject("btnFromCyrillicToLatin.Image")));
            this.btnFromCyrillicToLatin.Label = "Кириллдан Лотинга";
            this.btnFromCyrillicToLatin.Name = "btnFromCyrillicToLatin";
            this.btnFromCyrillicToLatin.OfficeImageId = "Translate";
            this.btnFromCyrillicToLatin.ScreenTip = "Кириллдан Лотинга (Ctrl + Alt + Shift + K)";
            this.btnFromCyrillicToLatin.ShowImage = true;
            this.btnFromCyrillicToLatin.SuperTip = "Танланган матн ёки бутун ҳужжатни бир босишда лотинга ўтказади.\nТезкор ишлатиш: C" +
    "trl + Alt + Shift + K.";
            this.btnFromCyrillicToLatin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnFromCyrillicToLatin_Click);
            // 
            // btnTransExceptions
            // 
            this.btnTransExceptions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnTransExceptions.Image = ((System.Drawing.Image)(resources.GetObject("btnTransExceptions.Image")));
            this.btnTransExceptions.Label = "Истиснолар";
            this.btnTransExceptions.Name = "btnTransExceptions";
            this.btnTransExceptions.OfficeImageId = "Dictionary";
            this.btnTransExceptions.ScreenTip = "Истисноларни бошқариш (Ctrl + Alt + X)";
            this.btnTransExceptions.ShowImage = true;
            this.btnTransExceptions.SuperTip = "Транслитерация учун истисно сўзлар рўйхатини очади.\nИсмлар, брендлар ва махсус ат" +
    "амаларни қўшиш ёки таҳрирлаш учун.";
            this.btnTransExceptions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnTransExceptions_Click);
            // 
            // groupTools
            // 
            this.groupTools.Items.Add(this.btnCleanupSpaces);
            this.groupTools.Items.Add(this.btnSwitchScript);
            this.groupTools.Items.Add(this.btnSetFontTNR);
            this.groupTools.Label = "Воситалар";
            this.groupTools.Name = "groupTools";
            // 
            // btnCleanupSpaces
            // 
            this.btnCleanupSpaces.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnCleanupSpaces.Label = "Бўшлиқларни тозалаш";
            this.btnCleanupSpaces.Name = "btnCleanupSpaces";
            this.btnCleanupSpaces.OfficeImageId = "ClearFormatting";
            this.btnCleanupSpaces.ScreenTip = "Бўшлиқларни тозалаш  (Ctrl + Alt + Shift + B)";
            this.btnCleanupSpaces.ShowImage = true;
            this.btnCleanupSpaces.SuperTip = "Ортиқча бўшлиқлар, икки марта босилган Enter ва бошқа формат хатоларини тозалайди" +
    ".\n(Матнни тартибга солиш учун.)";
            this.btnCleanupSpaces.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnCleanupSpaces_Click);
            // 
            // btnSwitchScript
            // 
            this.btnSwitchScript.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.btnSwitchScript.Image = ((System.Drawing.Image)(resources.GetObject("btnSwitchScript.Image")));
            this.btnSwitchScript.Label = "Ёзувни алмаштириш";
            this.btnSwitchScript.Name = "btnSwitchScript";
            this.btnSwitchScript.OfficeImageId = "InsertSymbol";
            this.btnSwitchScript.ScreenTip = "Ёзувни алмаштириш  (Ctrl + Alt + U)";
            this.btnSwitchScript.ShowImage = true;
            this.btnSwitchScript.SuperTip = "Ҳужжатнинг ҳозирги ёзувини аниқлаб, тескарисига (лотин ↔ кирилл) ўтказади.\n(Арала" +
    "ш ёки номаълум ёзувни тез ўзгартириш учун.)";
            this.btnSwitchScript.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSwitchScript_Click);
            // 
            // btnSetFontTNR
            // 
            this.btnSetFontTNR.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnSetFontTNR.Label = "TNR ўрнатиш";
            this.btnSetFontTNR.Name = "btnSetFontTNR";
            this.btnSetFontTNR.OfficeImageId = "FontDialog";
            this.btnSetFontTNR.ScreenTip = "TNR ўрнатиш  (Ctrl + Alt + Shift + N)";
            this.btnSetFontTNR.ShowImage = true;
            this.btnSetFontTNR.SuperTip = "Танланган матн ёки бутун ҳужжат шрифтини Times New Roman га ўрнатади.\n(Имло текши" +
    "рувининг тўғри ишлаши учун тавсия этилади.)";
            this.btnSetFontTNR.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSetFontTNR_Click);
            // 
            // groupExplanations
            // 
            this.groupExplanations.Items.Add(this.btnDefinitions);
            this.groupExplanations.Items.Add(this.btnExportErrors);
            this.groupExplanations.Label = "Изоҳлар";
            this.groupExplanations.Name = "groupExplanations";
            // 
            // btnDefinitions
            // 
            this.btnDefinitions.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnDefinitions.Label = "Изоҳ";
            this.btnDefinitions.Name = "btnDefinitions";
            this.btnDefinitions.OfficeImageId = "Thesaurus";
            this.btnDefinitions.ScreenTip = "Изоҳ  (Ctrl + Alt + Shift + Z)";
            this.btnDefinitions.ShowImage = true;
            this.btnDefinitions.SuperTip = "Танланган сўзнинг маъноси, имло қоидаси ёки изоҳини кўрсатади.\n(Сўзнинг тўғри ёзи" +
    "лишини тушуниш учун.)";
            this.btnDefinitions.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnDefinitions_Click);
            // 
            // btnExportErrors
            // 
            this.btnExportErrors.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnExportErrors.Label = "Хатоларни экспорт қилиш";
            this.btnExportErrors.Name = "btnExportErrors";
            this.btnExportErrors.OfficeImageId = "ExportExcel";
            this.btnExportErrors.ScreenTip = "Хатоларни экспорт қилиш  (Ctrl + Alt + Shift + O)";
            this.btnExportErrors.ShowImage = true;
            this.btnExportErrors.SuperTip = "Топилган хатолар рўйхатини янги ҳужжатга ёки файлга сақлайди.\n(Ҳисобот тайёрлаш ё" +
    "ки бошқалар билан бўлиш учун.)";
            this.btnExportErrors.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnExportErrors_Click);
            // 
            // groupSpecialChars
            // 
            this.groupSpecialChars.Items.Add(this.btnSpecialChar1);
            this.groupSpecialChars.Items.Add(this.btnSpecialChar2);
            this.groupSpecialChars.Items.Add(this.btnSpecialCharAll);
            this.groupSpecialChars.Items.Add(this.lblScriptIndicator);
            this.groupSpecialChars.Label = "Махсус белгилар";
            this.groupSpecialChars.Name = "groupSpecialChars";
            // 
            // btnSpecialChar1
            // 
            this.btnSpecialChar1.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnSpecialChar1.Image = ((System.Drawing.Image)(resources.GetObject("btnSpecialChar1.Image")));
            this.btnSpecialChar1.Label = "Тутуқ белгиси";
            this.btnSpecialChar1.Name = "btnSpecialChar1";
            this.btnSpecialChar1.ScreenTip = "(Ctrl + Alt + Shift + 1)";
            this.btnSpecialChar1.ShowImage = true;
            this.btnSpecialChar1.SuperTip = "Тутуқ белгисини жорий курсор жойига қўяди.\nТезкор киритиш учун Ctrl + Alt + Shift" +
    " + 1 ни босинг.";
            this.btnSpecialChar1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSpecialChar1_Click);
            // 
            // btnSpecialChar2
            // 
            this.btnSpecialChar2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnSpecialChar2.Image = ((System.Drawing.Image)(resources.GetObject("btnSpecialChar2.Image")));
            this.btnSpecialChar2.Label = "Тутуқ белгиси";
            this.btnSpecialChar2.Name = "btnSpecialChar2";
            this.btnSpecialChar2.ScreenTip = "(Ctrl + Alt + Shift + 2)";
            this.btnSpecialChar2.ShowImage = true;
            this.btnSpecialChar2.SuperTip = "Тутуқ белгисини жорий курсор жойига қўяди.\nТезкор киритиш учун Ctrl + Alt + Shift" +
    " + 2 ни босинг.";
            this.btnSpecialChar2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSpecialChar2_Click);
            // 
            // btnSpecialCharAll
            // 
            this.btnSpecialCharAll.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnSpecialCharAll.Image = ((System.Drawing.Image)(resources.GetObject("btnSpecialCharAll.Image")));
            this.btnSpecialCharAll.Label = "Тўғри тутуқ белгиси қўйиш";
            this.btnSpecialCharAll.Name = "btnSpecialCharAll";
            this.btnSpecialCharAll.ScreenTip = "Тўғри тутуқ белгиси қўйиш  (Ctrl + Alt + Shift + F)";
            this.btnSpecialCharAll.ShowImage = true;
            this.btnSpecialCharAll.SuperTip = "Матндаги нотўғри тутуқ белгиларини тўғри форматга келтиради.\nТезкор ишлатиш учун " +
    "Ctrl + Alt + Shift + F ни босинг.";
            this.btnSpecialCharAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnSpecialCharAll_Click);
            // 
            // lblScriptIndicator
            // 
            this.lblScriptIndicator.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.lblScriptIndicator.Label = "—";
            this.lblScriptIndicator.Name = "lblScriptIndicator";
            this.lblScriptIndicator.ScreenTip = "Ҳозирги ёзув тури";
            this.lblScriptIndicator.ShowImage = true;
            this.lblScriptIndicator.SuperTip = "Ҳужжатдаги матннинг ҳозирги ёзув турини кўрсатади (Кирилл ёки Лотин).";
            // 
            // groupInfo
            // 
            this.groupInfo.Items.Add(this.btnAppInfo);
            this.groupInfo.Label = "Маълумот";
            this.groupInfo.Name = "groupInfo";
            // 
            // btnAppInfo
            // 
            this.btnAppInfo.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeRegular;
            this.btnAppInfo.Label = "Маълумот";
            this.btnAppInfo.Name = "btnAppInfo";
            this.btnAppInfo.OfficeImageId = "Info";
            this.btnAppInfo.ScreenTip = "Дастур ҳақида маълумот";
            this.btnAppInfo.ShowImage = true;
            this.btnAppInfo.SuperTip = "Дастур ҳақида, қўлланма, махфиялик сиёсати ва қўллаб-қувватлаш маълумотларини кўр" +
    "сатади.";
            this.btnAppInfo.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btnAppInfo_Click);
            // 
            // UzbekOrfoRibbon
            // 
            this.Name = "UzbekOrfoRibbon";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tabUzbekOrfo);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.UzbekOrfoRibbon_Load);
            this.tabUzbekOrfo.ResumeLayout(false);
            this.tabUzbekOrfo.PerformLayout();
            this.groupCheck.ResumeLayout(false);
            this.groupCheck.PerformLayout();
            this.groupFix.ResumeLayout(false);
            this.groupFix.PerformLayout();
            this.groupDictionary.ResumeLayout(false);
            this.groupDictionary.PerformLayout();
            this.groupTransliteration.ResumeLayout(false);
            this.groupTransliteration.PerformLayout();
            this.groupTools.ResumeLayout(false);
            this.groupTools.PerformLayout();
            this.groupExplanations.ResumeLayout(false);
            this.groupExplanations.PerformLayout();
            this.groupSpecialChars.ResumeLayout(false);
            this.groupSpecialChars.PerformLayout();
            this.groupInfo.ResumeLayout(false);
            this.groupInfo.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private Microsoft.Office.Tools.Ribbon.RibbonTab tabUzbekOrfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupCheck;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupFix;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupDictionary;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTransliteration;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupTools;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupExplanations;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupSpecialChars;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup groupInfo;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCheckSpelling;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnViewErrors;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnErrorList;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSuggestions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnReplaceAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddToDict;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnEditDictionary;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAddNewWords;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton lblScriptIndicator;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFromLatinToCyrillic;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnFromCyrillicToLatin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnTransExceptions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnCleanupSpaces;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSwitchScript;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSetFontTNR;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnDefinitions;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnExportErrors;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSpecialChar1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSpecialChar2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnSpecialCharAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleAutoCorrect;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btnAppInfo;
    }

    partial class ThisRibbonCollection
    {
        internal UzbekOrfoRibbon UzbekOrfoRibbon
        {
            get { return this.GetRibbon<UzbekOrfoRibbon>(); }
        }
    }
}
