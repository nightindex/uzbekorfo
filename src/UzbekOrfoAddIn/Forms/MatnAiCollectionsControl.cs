using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using UzbekOrfoAddIn.Prediction;
using UzbekOrfoAddIn.UI;
using UzbekOrfoAddIn.UI.Controls;

namespace UzbekOrfoAddIn.Forms
{
    public sealed class MatnAiCollectionsControl : UserControl, IMatnAiSettingsWork
    {
        private readonly CollectionStore _store;
        private readonly Action _changed;
        private readonly Func<string[]> _getActive;
        private readonly Action<string[]> _setActive;
        private readonly Action<string> _forgetLearning;
        private readonly CollectionImporter _importer = new CollectionImporter();
        private readonly CheckedListBox _collections = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = false, IntegralHeight = false, BorderStyle = BorderStyle.None, HorizontalScrollbar = true };
        private readonly ListBox _files = new ListBox { Dock = DockStyle.Fill, SelectionMode = SelectionMode.MultiExtended, IntegralHeight = false, HorizontalScrollbar = true, BorderStyle = BorderStyle.None };
        private readonly TextBox _name = new TextBox { Dock = DockStyle.Fill, MaxLength = 100 };
        private readonly CheckBox _recursive = new CheckBox { Text = "Ички папкаларни ҳам қўшиш", AutoSize = true };
        private readonly Label _status = new Label { AutoSize = true, Dock = DockStyle.Fill, UseMnemonic = false };
        private readonly ProgressBar _progress = new ProgressBar { Dock = DockStyle.Fill, Style = ProgressBarStyle.Continuous };
        private readonly FlowLayoutPanel _actions = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
        private readonly ModernButton _cancel;
        private readonly List<string> _pending = new List<string>();
        private readonly Dictionary<string, bool> _pendingLocations = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);
        private List<PredictionCollection> _items = new List<PredictionCollection>();
        private CancellationTokenSource _cancellation;
        private Task _work = Task.CompletedTask;
        private bool _updating;
        private bool _layingOut;
        private bool _layoutQueued;
        public bool IsBusy { get; private set; }
        public event EventHandler BusyChanged;
        private PredictionCollection Selected => _collections.SelectedIndex < 0 ? null : _items[_collections.SelectedIndex];

        public MatnAiCollectionsControl(CollectionStore store, Action changed, Func<string[]> getActiveIds, Action<string[]> setActiveIds,
            Action<string> forgetLearning = null)
        {
            _store = store; _changed = changed; _getActive = getActiveIds; _setActive = setActiveIds;
            _forgetLearning = forgetLearning;
            AutoScaleMode = AutoScaleMode.None;
            Dock = DockStyle.Fill; AllowDrop = true; AutoScroll = true; BackColor = ThemeManager.Background;
            ForeColor = ThemeManager.TextPrimary; Font = ThemeManager.FontBase;
            // Keep list viewports usable at high DPI/narrow widths; scroll the page instead of collapsing them.
            var root = new Panel();
            var heading = new Label { Text = "Ҳужжатларингиздан ўрганинг", Font = ThemeManager.FontXLBold, AutoSize = false };
            var collectionHeading = new Label { Text = "Шу ҳужжат учун тўпламлар", Font = ThemeManager.FontBaseBold, AutoSize = false };
            var sourceHeading = new Label { Text = "Манбалар ва ўрганиш натижаси", Font = ThemeManager.FontBaseBold, AutoSize = false };
            root.Controls.Add(heading); root.Controls.Add(collectionHeading); root.Controls.Add(sourceHeading);
            var introduction = new Label { Text = "Белгиланган шахсий тўпламлар фақат шу очиқ ҳужжатда ишлайди. Тўплам амаллари дарҳол сақланади.", AutoSize = false };
            root.Controls.Add(introduction);
            root.Controls.Add(_collections);
            var names = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2 };
            names.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); names.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            names.Controls.Add(new Label { Text = "Тўплам номи", AutoSize = true, Anchor = AnchorStyles.Left }, 0, 0); names.Controls.Add(_name, 1, 0);
            root.Controls.Add(names);
            var options = new FlowLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, WrapContents = true };
            options.Controls.Add(_recursive);
            options.Controls.Add(new Label { Text = "DOCX, DOC, TXT • Файл ёки папкани бу ерга ташланг", AutoSize = true, Margin = new Padding(8) });
            root.Controls.Add(options);
            AddAction("Янги тўплам", () => { _collections.ClearSelected(); _name.Text = ""; _pending.Clear(); _pendingLocations.Clear(); _files.Items.Clear(); return Task.CompletedTask; });
            AddAction("Файл қўшиш", PickFiles);
            AddAction("Папка қўшиш", PickFolder);
            AddAction("Ўрганиш", StartImport, ModernButton.ButtonStyle.Primary);
            AddAction("Янгилаш", RefreshCollection);
            AddAction("Номини ўзгартириш", Rename);
            AddAction("Манбани олиб ташлаш", RemoveSources);
            AddAction("Тўпламни ўчириш", DeleteCollection, ModernButton.ButtonStyle.Danger);
            root.Controls.Add(_actions); root.Controls.Add(_files); root.Controls.Add(_progress);
            var footer = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 2 };
            footer.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100)); footer.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            _cancel = MakeButton("Тўхтатиш"); _cancel.Enabled = false; _cancel.Click += (s, e) => _cancellation?.Cancel();
            footer.Controls.Add(_status, 0, 0); footer.Controls.Add(_cancel, 1, 0); root.Controls.Add(footer);
            foreach (Control row in root.Controls) { row.Dock = DockStyle.None; row.AutoSize = false; }
            Controls.Add(root);
            Action arrange = () =>
            {
                if (_layingOut) return;
                _layingOut = true;
                try
                {
                    int dpi = (FindForm() as DpiForm)?.LayoutDpi ?? 96;
                    Func<int, int> px = value => (int)Math.Round(value * dpi / 96.0);
                    int gap = px(8), width = Math.Max(px(100), ClientSize.Width - SystemInformation.VerticalScrollBarWidth);
                    int rowWidth = Math.Max(px(80), width - 2 * gap), y = gap;
                    Action<Control, int> row = (child, height) => { child.Bounds = new Rectangle(gap, y, rowWidth, height); y += height + gap; };
                    row(heading, heading.GetPreferredSize(new Size(rowWidth, 0)).Height + px(8));
                    row(introduction, introduction.GetPreferredSize(new Size(rowWidth, 0)).Height + px(4));
                    row(collectionHeading, collectionHeading.GetPreferredSize(new Size(rowWidth, 0)).Height + px(4));
                    row(_collections, px(120));
                    row(names, px(34));
                    row(options, options.GetPreferredSize(new Size(rowWidth, 0)).Height);
                    int columns = Math.Max(1, rowWidth / px(160));
                    row(_actions, ((_actions.Controls.Count + columns - 1) / columns) * px(40));
                    row(sourceHeading, sourceHeading.GetPreferredSize(new Size(rowWidth, 0)).Height + px(4));
                    row(_files, px(160));
                    row(_progress, px(10));
                    row(footer, Math.Max(px(42), _status.GetPreferredSize(new Size(Math.Max(px(80), rowWidth - _cancel.Width - gap), 0)).Height));
                    root.Bounds = new Rectangle(0, AutoScrollPosition.Y, width, y);
                }
                finally { _layingOut = false; }
            };
            Layout += (s, e) => arrange();
            root.Layout += (s, e) => arrange();
            _status.TextChanged += (s, e) => arrange();
            // DpiForm scales every descendant before native controls finish their
            // own handle/font layout. Reconcile once after that pass has completed.
            EventHandler queueLayout = (s, e) =>
            {
                if (_layoutQueued || _layingOut || !IsHandleCreated || IsDisposed) return;
                _layoutQueued = true;
                BeginInvoke((Action)(() =>
                {
                    try { if (!IsDisposed) arrange(); }
                    finally { _layoutQueued = false; }
                }));
            };
            Layout += (s, e) => queueLayout(s, e);
            VisibleChanged += queueLayout;
            HandleCreated += queueLayout;
            foreach (Control control in new Control[] { _files, _collections, _name })
            { control.BackColor = ThemeManager.SurfaceElevated; control.ForeColor = ThemeManager.TextPrimary; }
            _collections.AccessibleName = "Шу ҳужжат учун тўпламлар"; _files.AccessibleName = "Манбалар ва ўрганиш учун танланган файллар";
            _name.AccessibleName = "Тўплам номи";
            _collections.SelectedIndexChanged += (s, e) => { if (!_updating) ShowSources(); };
            _collections.ItemCheck += CheckCollection;
            DragEnter += (s, e) => e.Effect = !IsBusy && e.Data.GetDataPresent(DataFormats.FileDrop) ? DragDropEffects.Copy : DragDropEffects.None;
            DragDrop += async (s, e) => await Run(() => AddPaths((string[])e.Data.GetData(DataFormats.FileDrop)));
            Load += async (s, e) => await Run(Reload);
        }

        private void AddAction(string title, Func<Task> action, ModernButton.ButtonStyle style = ModernButton.ButtonStyle.Secondary)
        {
            var button = MakeButton(title); button.Style = style;
            button.Click += async (s, e) => await Run(action); _actions.Controls.Add(button);
        }
        private static ModernButton MakeButton(string text) => new ModernButton { Text = text, AccessibleName = text,
            Style = ModernButton.ButtonStyle.Secondary, Size = new Size(154, 34), Margin = new Padding(3), Font = ThemeManager.FontBase };
        private async Task Run(Func<Task> action)
        {
            if (IsBusy) return;
            IsBusy = true; _actions.Enabled = _collections.Enabled = _files.Enabled = _name.Enabled = _recursive.Enabled = false;
            _cancel.Enabled = true; _cancellation = new CancellationTokenSource(); BusyChanged?.Invoke(this, EventArgs.Empty);
            try { _work = action(); await _work; }
            catch (OperationCanceledException) { _status.Text = "Тўхтатилди. Аввалги тўплам сақланди."; }
            catch (Exception ex) { _status.Text = "Амал бажарилмади."; ModernMessageBox.Error(ex.Message, "MatnAI"); }
            finally
            {
                _cancellation.Dispose(); _cancellation = null; IsBusy = false;
                _actions.Enabled = _collections.Enabled = _files.Enabled = _name.Enabled = _recursive.Enabled = true;
                _cancel.Enabled = false; BusyChanged?.Invoke(this, EventArgs.Empty);
                PerformLayout();
            }
        }
        public async Task CancelAndWaitAsync()
        {
            _cancellation?.Cancel();
            try { await _work; } catch (OperationCanceledException) { } catch { }
        }
        private async Task Reload()
        {
            string selected = Selected?.Id;
            var rows = await Task.Run(() => _store.LoadAll().Select(collection =>
            {
                int words = collection.Sources.Sum(s => s.WordCount);
                int phrases = collection.Sources.SelectMany(s => s.Sequences.Keys).Where(k => k.Contains(" ")).Distinct().Count();
                // Built-ins are read-only here: keep metadata, not their training
                // payload alongside the live engine. Never trim editable collections.
                var item = collection.BuiltIn ? new PredictionCollection
                {
                    Id = collection.Id, Name = collection.Name, BuiltIn = true, UpdatedUtc = collection.UpdatedUtc,
                    Sources = collection.Sources.Select(s => new PredictionSource
                    { Path = s.Path, SourceUrl = s.SourceUrl, Hash = s.Hash, WordCount = s.WordCount }).ToList()
                } : collection;
                return new { Item = item, Words = words, Phrases = phrases };
            }).ToList());
            var items = rows.Select(r => r.Item).ToList();
            _updating = true;
            try
            {
                _items = items; _collections.Items.Clear();
                var active = new HashSet<string>(_getActive());
                foreach (var row in rows)
                {
                    var collection = row.Item;
                    _collections.Items.Add(collection.Name + (collection.BuiltIn ? " [ҳуқуқий база]" : "") +
                        " — " + collection.Sources.Count + " файл, " + row.Words + " сўз, " + row.Phrases + " ибора; " + collection.UpdatedUtc.ToLocalTime().ToString("g"),
                        collection.BuiltIn || active.Contains(collection.Id));
                }
                _collections.SelectedIndex = items.FindIndex(c => c.Id == selected);
            }
            finally { _updating = false; }
            ShowSources();
            if (_store.LoadErrors.Count > 0) _status.Text = "Айрим тўпламлар очилмади: " + string.Join("; ", _store.LoadErrors);
        }
        private void CheckCollection(object sender, ItemCheckEventArgs e)
        {
            if (_updating) return;
            if (_items[e.Index].BuiltIn) { e.NewValue = CheckState.Checked; return; }
            var ids = _items.Where((item, index) => !item.BuiltIn &&
                (index == e.Index ? e.NewValue == CheckState.Checked : _collections.GetItemChecked(index))).Select(c => c.Id).ToArray();
            _setActive(ids);
        }
        private void ShowSources()
        {
            _pending.Clear(); _pendingLocations.Clear(); _files.Items.Clear();
            if (Selected != null) { _name.Text = Selected.Name; foreach (var source in Selected.Sources) _files.Items.Add(source.Path); }
            _status.Text = "Шахсий тўпламларда сўз ва иборалар шу компьютерда шифрлаб сақланади. Манба файллар ўзгармайди.";
        }
        private Task PickFiles()
        {
            using (var dialog = new OpenFileDialog { Multiselect = true, Filter = "Word ва матн|*.docx;*.doc;*.txt" })
                return dialog.ShowDialog(this) == DialogResult.OK ? AddPaths(dialog.FileNames) : Task.CompletedTask;
        }
        private Task PickFolder()
        {
            using (var dialog = new FolderBrowserDialog { Description = "Ўрганиш учун папкани танланг" })
                return dialog.ShowDialog(this) == DialogResult.OK ? AddPaths(new[] { dialog.SelectedPath }) : Task.CompletedTask;
        }
        private async Task AddPaths(string[] paths)
        {
            if (Selected?.BuiltIn == true) { _collections.ClearSelected(); _name.Text = ""; }
            bool recursive = _recursive.Checked;
            var token = _cancellation.Token;
            var found = await Task.Run(() => CollectionImporter.Discover(paths, recursive, token));
            foreach (string path in paths) _pendingLocations[Path.GetFullPath(path)] = recursive;
            if (Selected != null) found = found.Where(p => !Selected.ExcludedSources.Contains(p, StringComparer.OrdinalIgnoreCase) || paths.Contains(p, StringComparer.OrdinalIgnoreCase)).ToArray();
            _pending.AddRange(found.Except(_pending, StringComparer.OrdinalIgnoreCase));
            _files.Items.Clear(); _files.Items.AddRange(_pending.ToArray());
            _status.Text = _pending.Count + " файл танланди. Рўйхатни текшириб, «Ўрганиш» тугмасини босинг.";
        }
        private Task StartImport()
        {
            if (string.IsNullOrWhiteSpace(_name.Text)) throw new InvalidOperationException("Тўплам номини киритинг.");
            if (_pending.Count == 0) throw new InvalidOperationException("Аввал файл ёки папка танланг.");
            var collection = Selected;
            if (collection == null || collection.BuiltIn) collection = new PredictionCollection { Name = _name.Text.Trim() };
            return Import(collection, _pending.ToArray());
        }
        private async Task RefreshCollection()
        {
            var collection = Editable(); var token = _cancellation.Token;
            var discovered = await Task.Run(() => collection.ImportLocations.Where(p => Directory.Exists(p.Key))
                .SelectMany(p => CollectionImporter.Discover(new[] { p.Key }, p.Value, token)).ToArray(), token);
            await Import(collection, collection.Sources.Select(s => s.Path).Concat(discovered)
                .Where(p => !collection.ExcludedSources.Contains(p, StringComparer.OrdinalIgnoreCase)).Distinct(StringComparer.OrdinalIgnoreCase).ToArray());
        }
        private async Task Import(PredictionCollection collection, string[] paths)
        {
            var progress = new Progress<CollectionImportProgress>(p =>
            {
                if (IsDisposed) return;
                _progress.Value = p.TotalFiles == 0 ? 0 : p.CompletedFiles * 100 / p.TotalFiles;
                _status.Text = p.CompletedFiles + "/" + p.TotalFiles + " — " + Path.GetFileName(p.Path);
            });
            var token = _cancellation.Token;
            var result = await _importer.ImportAsync(collection, paths, progress, token);
            foreach (var location in _pendingLocations) result.Collection.ImportLocations[location.Key] = location.Value;
            result.Collection.ExcludedSources.RemoveAll(p => _pendingLocations.ContainsKey(p));
            token.ThrowIfCancellationRequested();
            if (!result.Files.Any(f => f.Status == CollectionImportStatus.Imported || f.Status == CollectionImportStatus.Duplicate))
            {
                _files.Items.Clear(); _files.Items.AddRange(result.Files.Select(f => f.Path + " — " + f.Message).ToArray());
                _status.Text = "Ҳеч бир файл ўрганилмади. Аввалги тўплам сақланди."; return;
            }
            await Task.Run(() => _store.Save(result.Collection, cancellation: token));
            _changed(); await Reload();
            _files.Items.Clear(); _files.Items.AddRange(result.Files.Select(f => f.Path + " — " + f.Message).ToArray());
            _status.Text = "Ўрганилди: " + result.Files.Count(f => f.Status == CollectionImportStatus.Imported) +
                "; такрор: " + result.Files.Count(f => f.Status == CollectionImportStatus.Duplicate) +
                "; ўтказиб юборилди: " + result.Files.Count(f => f.Status != CollectionImportStatus.Imported && f.Status != CollectionImportStatus.Duplicate) +
                ". Тўпламни шу ҳужжат учун белгиланг.";
        }
        private PredictionCollection Editable()
        {
            if (Selected == null || Selected.BuiltIn) throw new InvalidOperationException("Шахсий тўпламни танланг.");
            return Selected;
        }
        private async Task Rename()
        {
            var collection = Editable();
            if (string.IsNullOrWhiteSpace(_name.Text)) throw new InvalidOperationException("Тўплам номини киритинг.");
            collection.Name = _name.Text.Trim(); await Task.Run(() => _store.Save(collection)); _changed(); await Reload();
        }
        private async Task RemoveSources()
        {
            var collection = Editable();
            var labels = _files.SelectedItems.Cast<string>().ToArray();
            var paths = collection.Sources.Where(s => labels.Any(label => label == s.Path || label.StartsWith(s.Path + " — ", StringComparison.Ordinal)))
                .Select(s => s.Path).ToArray();
            if (paths.Length == 0) return;
            if (!ModernMessageBox.Confirm("Танланган манбалардан ўрганилган маълумотлар олиб ташлансинми? Асл файллар ўчирилмайди.", "MatnAI", "Олиб ташлаш", "Бекор қилиш")) return;
            collection.Sources.RemoveAll(s => paths.Contains(s.Path)); CollectionImporter.Recount(collection);
            collection.ExcludedSources.AddRange(paths.Except(collection.ExcludedSources, StringComparer.OrdinalIgnoreCase));
            await Task.Run(() => _store.Save(collection, discardPrevious: true)); _forgetLearning?.Invoke(collection.Id); _changed(); await Reload();
        }
        private async Task DeleteCollection()
        {
            var collection = Editable();
            if (!ModernMessageBox.Confirm("Тўплам ва унинг тиклаш нусхалари ўчирилсинми? Асл файллар ўчирилмайди.", "MatnAI", "Ўчириш", "Бекор қилиш")) return;
            await Task.Run(() => _store.Delete(collection.Id));
            _forgetLearning?.Invoke(collection.Id);
            _setActive(_getActive().Where(id => id != collection.Id).ToArray()); _changed(); await Reload();
        }
    }
}
