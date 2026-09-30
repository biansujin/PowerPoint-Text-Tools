using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using Media = System.Windows.Media;
using Extensibility;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

[assembly: ComVisible(true)]
[assembly: Guid("9A85B0D1-57B8-4B6A-8B54-17F8C0B45A10")]

namespace OfficeSmartFontPicker
{
    [ComVisible(true)]
    [Guid("A3D9E13C-5732-4E7E-B6C7-0A4E8D8395B1")]
    [ProgId("OfficeSmartFontPicker.Connect")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class Connect : IDTExtensibility2, Office.IRibbonExtensibility, Office.ICustomTaskPaneConsumer
    {
        internal static object ApplicationInstance;
        private Office.ICTPFactory _ctpFactory;
        private Office.CustomTaskPane _taskPane;
        private PowerPoint.Application _powerPointApp;

        public void OnConnection(object application, ext_ConnectMode connectMode, object addInInst, ref Array custom)
        {
            ApplicationInstance = application;
            try
            {
                _powerPointApp = application as PowerPoint.Application;
                if (_powerPointApp != null)
                    _powerPointApp.WindowSelectionChange += PowerPointSelectionChanged;
            }
            catch (Exception ex)
            {
                Log("Selection event hookup ERROR " + ex);
            }
            Log("OnConnection");
        }

        public void CTPFactoryAvailable(Office.ICTPFactory CTPFactoryInst)
        {
            _ctpFactory = CTPFactoryInst;
            Log("CTPFactoryAvailable");
        }

        public void OnDisconnection(ext_DisconnectMode removeMode, ref Array custom)
        {
            try
            {
                if (_powerPointApp != null)
                    _powerPointApp.WindowSelectionChange -= PowerPointSelectionChanged;
            }
            catch { }
            try { if (_taskPane != null) _taskPane.Visible = false; } catch { }
            _taskPane = null;
            _ctpFactory = null;
            _powerPointApp = null;
            ApplicationInstance = null;
            Log("OnDisconnection");
        }

        private void PowerPointSelectionChanged(PowerPoint.Selection selection)
        {
            try
            {
                var pane = FontPickerPane.CurrentInstance;
                if (pane != null)
                    pane.SyncFromPowerPointSelection(selection);
            }
            catch (Exception ex)
            {
                Log("Selection sync ERROR " + ex);
            }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string ribbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
              <ribbon><tabs><tab idMso='TabHome'>
              <group id='OfficeSmartFontPicker_Group' label='&#x6587;&#x5B57;&#x5DE5;&#x5177;' insertAfterMso='GroupFont'>
              <button id='OfficeSmartFontPicker_Open'
                      label='&#x6587;&#x5B57;&#x5DE5;&#x5177;'
                      size='large'
                      imageMso='FontDialog'
                      screentip='&#x6587;&#x5B57;&#x5DE5;&#x5177;'
                      supertip='&#x6253;&#x5F00;&#x53F3;&#x4FA7;&#x6587;&#x5B57;&#x4E0E;&#x6392;&#x7248;&#x5DE5;&#x5177;&#x680F;&#x3002;'
                      onAction='OnOpenPicker'/>
              </group></tab></tabs></ribbon>
            </customUI>";
        }

        public void OnOpenPicker(Office.IRibbonControl control)
        {
            try
            {
                Log("OnOpenPicker");
                EnsureTaskPane();
                if (_taskPane != null)
                    _taskPane.Visible = !_taskPane.Visible || true;
            }
            catch (Exception ex)
            {
                Log("OnOpenPicker ERROR " + ex);
                MessageBox.Show("无法打开字体侧边栏：\r\n" + ex.Message,
                    "\u6587\u5b57\u5de5\u5177", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void OnFitSlideTextBoxes(Office.IRibbonControl control)
        {
            try
            {
                int count = SlideTextBoxTools.FitCurrentSlide(ApplicationInstance);
                Log("Fit text boxes: " + count);
            }
            catch (Exception ex)
            {
                Log("Fit text boxes ERROR " + ex);
                MessageBox.Show("无法调整本页文本框：\r\n" + ex.Message,
                    "\u6587\u5b57\u5de5\u5177", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EnsureTaskPane()
        {
            if (_taskPane != null) return;
            if (_ctpFactory == null)
                throw new InvalidOperationException("PowerPoint 尚未提供侧边栏工厂，请重新启动 PowerPoint。");

            object window = Type.Missing;
            try
            {
                dynamic app = ApplicationInstance;
                if (app != null && app.ActiveWindow != null) window = app.ActiveWindow;
            }
            catch { }

            _taskPane = _ctpFactory.CreateCTP(
                "OfficeSmartFontPicker.FontPickerPane",
                "\u6587\u5b57\u5de5\u5177",
                window);

            _taskPane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionRight;
            _taskPane.Width = 380;
            _taskPane.Visible = true;
            try
            {
                if (FontPickerPane.CurrentInstance != null)
                    FontPickerPane.CurrentInstance.SyncFromCurrentSelection();
            }
            catch { }
            Log("TaskPane created");
        }

        internal static void Log(string message)
        {
            try
            {
                System.IO.File.AppendAllText(
                    @"E:\OfficeSmartFontPicker\debug.log",
                    DateTime.Now.ToString("s") + " " + message + "\r\n");
            }
            catch { }
        }
    }

    internal sealed class FontFaceInfo
    {
        public string Family;
        public string Face;
        public string ApplyName;
        public bool ExactRegistered;
        public bool Bold;
        public bool Italic;
        public int Weight;
        public HashSet<string> Aliases = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
        public string DisplayFace { get { return string.IsNullOrWhiteSpace(Face) ? "Regular" : Face; } }
        public override string ToString() { return DisplayFace; }
    }

    internal sealed class FontFamilyInfo
    {
        public string Name;
        public List<FontFaceInfo> Faces = new List<FontFaceInfo>();
        public HashSet<string> Aliases = new HashSet<string>(StringComparer.CurrentCultureIgnoreCase);
        public string SearchText;
    }

    internal static class FontCatalog
    {
        private static readonly List<string> RegisteredFontNames = LoadRegisteredFontNames();

        private static List<string> LoadRegisteredFontNames()
        {
            var names = new List<string>();
            try
            {
                var hives = new[] {
                    Microsoft.Win32.Registry.LocalMachine,
                    Microsoft.Win32.Registry.CurrentUser
                };
                foreach (var hive in hives)
                {
                    using (var key = hive.OpenSubKey(@"SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts"))
                    {
                        if (key == null) continue;
                        foreach (var valueName in key.GetValueNames())
                        {
                            var clean = (valueName ?? "").Trim();
                            int suffix = clean.LastIndexOf(" (", StringComparison.Ordinal);
                            if (suffix > 0) clean = clean.Substring(0, suffix).Trim();

                            // TTC / collection registrations can expose several face names
                            // in one registry value separated by ampersands.
                            foreach (var part in clean.Split(new[] { '&' }, StringSplitOptions.RemoveEmptyEntries))
                            {
                                var name = part.Trim();
                                if (!string.IsNullOrWhiteSpace(name)) names.Add(name);
                            }
                        }
                    }
                }
            }
            catch { }

            return names
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }

        private static string NormalizeFontKey(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            return new string(s.ToLowerInvariant()
                .Where(ch => char.IsLetterOrDigit(ch))
                .ToArray());
        }

        private static string ResolveApplyName(
            string familyName,
            IEnumerable<string> familyAliases,
            string faceName,
            IEnumerable<string> faceAliases,
            out bool exactRegistered)
        {
            exactRegistered = false;

            var families = new List<string>();
            families.Add(familyName);
            if (familyAliases != null) families.AddRange(familyAliases);
            families = families
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var faces = new List<string>();
            faces.Add(faceName);
            if (faceAliases != null) faces.AddRange(faceAliases);
            faces = faces
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            var orderedCandidates = new List<string>();

            // Most reliable: the WPF typographic family + the WPF face name.
            orderedCandidates.Add(familyName + " " + faceName);

            // Then try localized/Win32 family aliases with the SAME primary face.
            foreach (var fam in families)
                orderedCandidates.Add(fam + " " + faceName);

            // Only after that try alternate face aliases.
            foreach (var face in faces.Skip(1))
            {
                orderedCandidates.Add(familyName + " " + face);
                foreach (var fam in families.Skip(1))
                    orderedCandidates.Add(fam + " " + face);
            }

            bool regular = faceName.Equals("Regular", StringComparison.OrdinalIgnoreCase)
                        || faceName.Equals("Normal", StringComparison.OrdinalIgnoreCase)
                        || faceName.Equals("Roman", StringComparison.OrdinalIgnoreCase);

            if (regular)
            {
                orderedCandidates.Add(familyName + " Regular");
                orderedCandidates.Add(familyName);
                foreach (var fam in families) orderedCandidates.Add(fam);
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var candidate in orderedCandidates)
            {
                var key = NormalizeFontKey(candidate);
                if (string.IsNullOrEmpty(key) || !seen.Add(key)) continue;

                var registered = RegisteredFontNames.FirstOrDefault(
                    x => NormalizeFontKey(x).Equals(key, StringComparison.OrdinalIgnoreCase));

                if (!string.IsNullOrEmpty(registered))
                {
                    exactRegistered = true;
                    return registered;
                }
            }

            string fallback = regular ? familyName : familyName + " " + faceName;
            return fallback.Trim();
        }

        public static List<FontFamilyInfo> Load()
        {
            var families = new List<FontFamilyInfo>();

            foreach (var ff in Media.Fonts.SystemFontFamilies)
            {
                try
                {
                    var familyNames = ff.FamilyNames.Values
                        .Where(x => !string.IsNullOrWhiteSpace(x))
                        .Distinct(StringComparer.CurrentCultureIgnoreCase)
                        .ToList();

                    string familyName = familyNames.FirstOrDefault() ?? ff.Source;
                    if (string.IsNullOrWhiteSpace(familyName)) continue;

                    var info = new FontFamilyInfo { Name = familyName };
                    foreach (var n in familyNames) info.Aliases.Add(n);
                    info.Aliases.Add(ff.Source);

                    foreach (var tf in ff.GetTypefaces())
                    {
                        try
                        {
                            Media.GlyphTypeface gt;
                            if (!tf.TryGetGlyphTypeface(out gt)) continue;

                            var faceNames = gt.FaceNames.Values
                                .Where(x => !string.IsNullOrWhiteSpace(x))
                                .Distinct(StringComparer.CurrentCultureIgnoreCase)
                                .ToList();

                            string face = faceNames.FirstOrDefault() ?? tf.Weight.ToString();
                            var fi = new FontFaceInfo
                            {
                                Family = familyName,
                                Face = face,
                                Bold = tf.Weight >= System.Windows.FontWeights.SemiBold,
                                Italic = tf.Style == System.Windows.FontStyles.Italic
                                      || tf.Style == System.Windows.FontStyles.Oblique,
                                Weight = tf.Weight.ToOpenTypeWeight()
                            };

                            foreach (var n in familyNames) fi.Aliases.Add(n);
                            foreach (var n in faceNames) fi.Aliases.Add(n);
                            foreach (var n in gt.FamilyNames.Values)
                            {
                                fi.Aliases.Add(n);
                                info.Aliases.Add(n);
                            }
                            foreach (var n in gt.Win32FamilyNames.Values)
                            {
                                fi.Aliases.Add(n);
                                info.Aliases.Add(n);
                            }
                            foreach (var n in gt.Win32FaceNames.Values) fi.Aliases.Add(n);

                            bool exactRegistered;
                            fi.ApplyName = ResolveApplyName(
                                familyName,
                                familyNames.Concat(gt.Win32FamilyNames.Values).Concat(gt.FamilyNames.Values),
                                face,
                                faceNames.Concat(gt.Win32FaceNames.Values),
                                out exactRegistered);
                            fi.ExactRegistered = exactRegistered;
                            info.Faces.Add(fi);
                        }
                        catch { }
                    }

                    if (info.Faces.Count == 0)
                        info.Faces.Add(new FontFaceInfo
                        {
                            Family = familyName, Face = "Regular",
                            ApplyName = familyName, ExactRegistered = false, Weight = 400
                        });

                    info.Faces = info.Faces
                        .GroupBy(x => x.DisplayFace, StringComparer.CurrentCultureIgnoreCase)
                        .Select(g => g.OrderBy(x => x.Weight).First())
                        .OrderBy(x => x.Weight)
                        .ThenBy(x => x.DisplayFace)
                        .ToList();

                    info.SearchText = string.Join(" ", info.Aliases);
                    families.Add(info);
                }
                catch { }
            }

            return families
                .GroupBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase)
                .Select(g =>
                {
                    var first = g.First();
                    foreach (var more in g.Skip(1))
                    {
                        foreach (var a in more.Aliases) first.Aliases.Add(a);
                        foreach (var f in more.Faces)
                            if (!first.Faces.Any(x => x.DisplayFace.Equals(
                                f.DisplayFace, StringComparison.CurrentCultureIgnoreCase)))
                                first.Faces.Add(f);
                    }
                    first.SearchText = string.Join(" ", first.Aliases);
                    return first;
                })
                .OrderBy(x => x.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();
        }
    }

    internal static class SlideTextBoxTools
    {
        public static int FitCurrentSlide(object application)
        {
            if (application == null)
                throw new InvalidOperationException("PowerPoint \u5c1a\u672a\u8fde\u63a5\u3002");

            dynamic app = application;
            dynamic window = app.ActiveWindow;
            if (window == null)
                throw new InvalidOperationException("\u5f53\u524d\u6ca1\u6709\u53ef\u7528\u7684 PowerPoint \u7a97\u53e3\u3002");

            dynamic slide = window.View.Slide;
            if (slide == null)
                throw new InvalidOperationException("\u5f53\u524d\u89c6\u56fe\u6ca1\u6709\u6d3b\u52a8\u5e7b\u706f\u7247\u3002");

            int count = 0;
            dynamic shapes = slide.Shapes;
            for (int i = 1; i <= shapes.Count; i++)
                count += FitShape(shapes.Item(i));

            return count;
        }

        private static int FitShape(dynamic shape)
        {
            try
            {
                if ((int)shape.Type == 6) // msoGroup
                {
                    int nested = 0;
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        nested += FitShape(items.Item(i));
                    return nested;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame != -1) return 0;
                if ((int)shape.TextFrame.HasText != -1) return 0;

                int shapeType = (int)shape.Type;

                // Real PowerPoint text containers.
                bool eligible = shapeType == 17 || shapeType == 14; // TextBox / Placeholder

                // Some third-party/generated PPTs store plain text as a transparent
                // AutoShape instead of msoTextBox. Treat those as text boxes too.
                if (!eligible && (shapeType == 1 || shapeType == 5))
                {
                    bool fillInvisible = true;
                    bool lineInvisible = true;

                    try
                    {
                        fillInvisible = (int)shape.Fill.Visible != -1
                            || (float)shape.Fill.Transparency >= 0.98f;
                    }
                    catch { }

                    try
                    {
                        lineInvisible = (int)shape.Line.Visible != -1
                            || (float)shape.Line.Transparency >= 0.98f;
                    }
                    catch { }

                    eligible = fillInvisible && lineInvisible;
                }

                if (!eligible)
                {
                    try
                    {
                        string skipped = Convert.ToString(shape.TextFrame.TextRange.Text);
                        skipped = (skipped ?? "").Replace("\r", " ").Replace("\n", " ").Trim();
                        if (skipped.Length > 70) skipped = skipped.Substring(0, 70);
                        Connect.Log("FIT SKIP type=" + shapeType + " text=" + skipped);
                    }
                    catch { }
                    return 0;
                }

                // Never resize tables as text boxes.
                try { if ((int)shape.HasTable == -1) return 0; } catch { }

                dynamic tf2 = shape.TextFrame2;
                dynamic tr = tf2.TextRange;

                string preview = "";
                try
                {
                    preview = Convert.ToString(tr.Text) ?? "";
                    preview = preview.Replace("\r", " ").Replace("\n", " ").Trim();
                    if (preview.Length > 70) preview = preview.Substring(0, 70);
                }
                catch { }

                float marginLeft = 0f, marginRight = 0f, marginTop = 0f, marginBottom = 0f;
                try { marginLeft = (float)tf2.MarginLeft; } catch { }
                try { marginRight = (float)tf2.MarginRight; } catch { }
                try { marginTop = (float)tf2.MarginTop; } catch { }
                try { marginBottom = (float)tf2.MarginBottom; } catch { }

                float currentWidth = (float)shape.Width;
                float currentHeight = (float)shape.Height;

                float beforeTextLeft = (float)tr.BoundLeft;
                float beforeTextTop = (float)tr.BoundTop;

                int visualLineCount = 1;
                try { visualLineCount = Math.Max(1, (int)tr.Lines.Count); } catch { }

                float contentWidth;
                float contentHeight;
                string method;

                if (visualLineCount == 1)
                {
                    // Important for titles: with WordWrap enabled PowerPoint can report
                    // the available text column as BoundWidth instead of glyph width.
                    // Temporarily disable wrapping only for the measurement, then restore.
                    int oldWrap = 0;
                    bool haveWrap = false;
                    try
                    {
                        oldWrap = (int)tf2.WordWrap;
                        haveWrap = true;
                        tf2.WordWrap = 0; // msoFalse
                    }
                    catch { }

                    try
                    {
                        contentWidth = Math.Max(1f, (float)tr.BoundWidth);
                        contentHeight = Math.Max(1f, (float)tr.BoundHeight);
                    }
                    finally
                    {
                        if (haveWrap)
                        {
                            try { tf2.WordWrap = oldWrap; } catch { }
                        }
                    }

                    method = "single";
                }
                else
                {
                    // Preserve the existing visual line breaks. Fit to the union of all
                    // rendered lines instead of the text column width.
                    float minLeft = float.MaxValue;
                    float minTop = float.MaxValue;
                    float maxRight = float.MinValue;
                    float maxBottom = float.MinValue;

                    try
                    {
                        int lineCount = (int)tr.Lines.Count;
                        for (int i = 1; i <= lineCount; i++)
                        {
                            dynamic line = tr.Lines(i, 1);
                            float l = (float)line.BoundLeft;
                            float t = (float)line.BoundTop;
                            float w = (float)line.BoundWidth;
                            float h = (float)line.BoundHeight;
                            if (w <= 0f || h <= 0f) continue;

                            minLeft = Math.Min(minLeft, l);
                            minTop = Math.Min(minTop, t);
                            maxRight = Math.Max(maxRight, l + w);
                            maxBottom = Math.Max(maxBottom, t + h);
                        }
                    }
                    catch { }

                    if (minLeft == float.MaxValue)
                    {
                        contentWidth = Math.Max(1f, (float)tr.BoundWidth);
                        contentHeight = Math.Max(1f, (float)tr.BoundHeight);
                    }
                    else
                    {
                        contentWidth = Math.Max(1f, maxRight - minLeft);
                        contentHeight = Math.Max(1f, maxBottom - minTop);
                    }

                    method = "multi";
                }

                float measuredWidth = Math.Max(1f, contentWidth + marginLeft + marginRight);
                float measuredHeight = Math.Max(1f, contentHeight + marginTop + marginBottom);

                // This tool is a "trim text box" command, not an auto-grow command.
                // Never make an existing frame larger; this also makes repeated clicks
                // completely stable.
                float targetWidth = Math.Min(currentWidth, measuredWidth);
                float targetHeight = Math.Min(currentHeight, measuredHeight);

                bool widthAlreadyFit = Math.Abs(currentWidth - targetWidth) < 0.30f;
                bool heightAlreadyFit = Math.Abs(currentHeight - targetHeight) < 0.30f;

                if (widthAlreadyFit && heightAlreadyFit)
                {
                    Connect.Log(
                        "FIT stable type=" + shapeType
                        + " method=" + method
                        + " size=" + currentWidth.ToString("0.0") + "x" + currentHeight.ToString("0.0")
                        + " text=" + preview);
                    return 1;
                }

                int oldLock = 0;
                try
                {
                    oldLock = (int)shape.LockAspectRatio;
                    shape.LockAspectRatio = 0;
                }
                catch { }

                try { tf2.AutoSize = 0; } catch { }

                if (!widthAlreadyFit) shape.Width = targetWidth;
                if (!heightAlreadyFit) shape.Height = targetHeight;

                // Resizing Center/Right/Middle/Bottom anchored text moves the glyphs.
                // Counter-shift the box so the text itself stays where it was.
                float afterTextLeft = (float)tr.BoundLeft;
                float afterTextTop = (float)tr.BoundTop;

                float dx = beforeTextLeft - afterTextLeft;
                float dy = beforeTextTop - afterTextTop;
                if (Math.Abs(dx) > 0.01f) shape.Left = (float)shape.Left + dx;
                if (Math.Abs(dy) > 0.01f) shape.Top = (float)shape.Top + dy;

                // One final correction for Office floating-point rounding.
                try
                {
                    float finalLeft = (float)tr.BoundLeft;
                    float finalTop = (float)tr.BoundTop;
                    float fdx = beforeTextLeft - finalLeft;
                    float fdy = beforeTextTop - finalTop;
                    if (Math.Abs(fdx) > 0.05f) shape.Left = (float)shape.Left + fdx;
                    if (Math.Abs(fdy) > 0.05f) shape.Top = (float)shape.Top + fdy;
                }
                catch { }

                try { shape.LockAspectRatio = oldLock; } catch { }

                Connect.Log(
                    "FIT type=" + shapeType
                    + " method=" + method
                    + " old=" + currentWidth.ToString("0.0") + "x" + currentHeight.ToString("0.0")
                    + " new=" + ((float)shape.Width).ToString("0.0") + "x" + ((float)shape.Height).ToString("0.0")
                    + " text=" + preview);

                return 1;
            }
            catch (Exception ex)
            {
                try { Connect.Log("FIT ERROR " + ex.Message); } catch { }
                return 0;
            }
        }
    }

    internal sealed class OfficeRedColorTable : ProfessionalColorTable
    {
        private static readonly Color Red = Color.FromArgb(192, 50, 45);
        private static readonly Color RedLight = Color.FromArgb(252, 235, 234);
        private static readonly Color RedPressed = Color.FromArgb(247, 221, 219);
        private static readonly Color Border = Color.FromArgb(221, 224, 229);

        public override Color ToolStripBorder { get { return Color.Transparent; } }
        public override Color ToolStripGradientBegin { get { return Color.Transparent; } }
        public override Color ToolStripGradientMiddle { get { return Color.Transparent; } }
        public override Color ToolStripGradientEnd { get { return Color.Transparent; } }

        public override Color ButtonSelectedHighlight { get { return RedLight; } }
        public override Color ButtonSelectedGradientBegin { get { return RedLight; } }
        public override Color ButtonSelectedGradientMiddle { get { return RedLight; } }
        public override Color ButtonSelectedGradientEnd { get { return RedLight; } }
        public override Color ButtonSelectedBorder { get { return Red; } }

        public override Color ButtonPressedHighlight { get { return RedPressed; } }
        public override Color ButtonPressedGradientBegin { get { return RedPressed; } }
        public override Color ButtonPressedGradientMiddle { get { return RedPressed; } }
        public override Color ButtonPressedGradientEnd { get { return RedPressed; } }
        public override Color ButtonPressedBorder { get { return Red; } }

        public override Color CheckBackground { get { return RedLight; } }
        public override Color CheckSelectedBackground { get { return RedLight; } }
        public override Color CheckPressedBackground { get { return RedPressed; } }

        public override Color SeparatorDark { get { return Border; } }
        public override Color SeparatorLight { get { return Color.White; } }
    }

    internal sealed class OfficeRedRenderer : ToolStripProfessionalRenderer
    {
        public OfficeRedRenderer() : base(new OfficeRedColorTable())
        {
            RoundedEdges = false;
        }
    }

    internal sealed class WeightSlider : Control
    {
        private int _minimum;
        private int _maximum;
        private int _value;
        private bool _dragging;
        private List<int> _labels = new List<int>();

        public event EventHandler ValueChanged;

        public int Minimum
        {
            get { return _minimum; }
            set { _minimum = value; if (_value < _minimum) Value = _minimum; Invalidate(); }
        }

        public int Maximum
        {
            get { return _maximum; }
            set { _maximum = Math.Max(_minimum, value); if (_value > _maximum) Value = _maximum; Invalidate(); }
        }

        public int Value
        {
            get { return _value; }
            set
            {
                int v = Math.Max(_minimum, Math.Min(_maximum, value));
                if (_value == v) return;
                _value = v;
                Invalidate();
                if (ValueChanged != null) ValueChanged(this, EventArgs.Empty);
            }
        }

        public int TickFrequency { get; set; }
        public int SmallChange { get; set; }
        public int LargeChange { get; set; }

        public Color AccentColor { get; set; }
        public Color TrackColor { get; set; }
        public Color LabelColor { get; set; }

        public WeightSlider()
        {
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                     ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            Height = 58;
            Minimum = 0;
            Maximum = 0;
            SmallChange = 1;
            LargeChange = 1;
            AccentColor = Color.FromArgb(192, 50, 45);
            TrackColor = Color.FromArgb(217, 220, 225);
            LabelColor = Color.FromArgb(120, 125, 132);
            Cursor = Cursors.Hand;
            TabStop = false;
        }

        public void SetLabels(IEnumerable<int> values)
        {
            _labels = values == null ? new List<int>() : values.ToList();
            Invalidate();
        }

        private float ThumbX
        {
            get
            {
                int left = 12;
                int right = Math.Max(left + 1, ClientSize.Width - 12);
                if (_maximum <= _minimum) return left;
                float t = (float)(_value - _minimum) / (float)(_maximum - _minimum);
                return left + (right - left) * t;
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            var g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            int left = 12;
            int right = Math.Max(left + 1, ClientSize.Width - 12);
            int y = 18;
            float tx = ThumbX;

            using (var pen = new Pen(TrackColor, 3f))
            {
                pen.StartCap = pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                g.DrawLine(pen, left, y, right, y);
            }

            using (var pen = new Pen(AccentColor, 3f))
            {
                pen.StartCap = pen.EndCap = System.Drawing.Drawing2D.LineCap.Round;
                g.DrawLine(pen, left, y, tx, y);
            }

            using (var shadow = new SolidBrush(Color.FromArgb(28, 0, 0, 0)))
                g.FillEllipse(shadow, tx - 7, y - 5, 14, 14);

            using (var fill = new SolidBrush(AccentColor))
                g.FillEllipse(fill, tx - 6, y - 6, 12, 12);

            if (!Enabled)
            {
                using (var veil = new SolidBrush(Color.FromArgb(120, 244, 245, 247)))
                    g.FillRectangle(veil, ClientRectangle);
            }

            if (_labels.Count > 0)
            {
                using (var f = new Font("Segoe UI", 7.5f))
                using (var b = new SolidBrush(LabelColor))
                {
                    int maxLabels = Math.Min(_labels.Count, 7);
                    if (_labels.Count <= 7)
                    {
                        for (int i = 0; i < _labels.Count; i++)
                            DrawLabel(g, f, b, _labels[i].ToString(), i, _labels.Count, left, right);
                    }
                    else
                    {
                        int[] idx = new[] { 0, _labels.Count / 2, _labels.Count - 1 };
                        foreach (int i in idx)
                            DrawLabel(g, f, b, _labels[i].ToString(), i, _labels.Count, left, right);
                    }
                }
            }
        }

        private void DrawLabel(Graphics g, Font f, Brush b, string text, int index, int count, int left, int right)
        {
            float x = count <= 1 ? left : left + (right - left) * index / (float)(count - 1);
            SizeF sz = g.MeasureString(text, f);
            g.DrawString(text, f, b, x - sz.Width / 2f, 34f);
        }

        private void SetFromX(int x)
        {
            if (!Enabled || _maximum <= _minimum) return;
            int left = 12;
            int right = Math.Max(left + 1, ClientSize.Width - 12);
            float t = Math.Max(0f, Math.Min(1f, (x - left) / (float)(right - left)));
            int v = _minimum + (int)Math.Round(t * (_maximum - _minimum));
            Value = v;
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (!Enabled || e.Button != MouseButtons.Left) return;
            _dragging = true;
            Capture = true;
            SetFromX(e.X);
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (_dragging) SetFromX(e.X);
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            _dragging = false;
            Capture = false;
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (!Enabled) return;
            Value += e.Delta > 0 ? -1 : 1;
        }
    }

    internal sealed class ScrubSpinner : UserControl
    {
        private readonly TextBox _text;
        private readonly SpinnerArrows _arrows;
        private decimal _value;
        private decimal _minimum = 0M;
        private decimal _maximum = 100M;
        private decimal _increment = 1M;
        private int _decimalPlaces;
        private bool _internalUpdate;
        private bool _scrubbing;
        private int _scrubStartScreenY;
        private decimal _scrubStartValue;
        private bool _hover;
        private bool _focusWithin;

        public event EventHandler ValueChanged;

        public ScrubSpinner()
        {
            Height = 30;
            BackColor = Color.White;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                     ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);

            _arrows = new SpinnerArrows
            {
                Dock = DockStyle.Right,
                Width = 24
            };
            _arrows.StepUp += (s, e) => Step(1);
            _arrows.StepDown += (s, e) => Step(-1);

            _text = new TextBox
            {
                BorderStyle = BorderStyle.None,
                TextAlign = HorizontalAlignment.Right,
                Font = new Font("Segoe UI", 9F),
                BackColor = Color.White,
                ForeColor = Color.FromArgb(38, 42, 48),
                Left = 6,
                Top = 6,
                Anchor = AnchorStyles.Left | AnchorStyles.Right | AnchorStyles.Top
            };
            _text.KeyDown += TextKeyDown;
            _text.Leave += (s, e) => { _focusWithin = false; CommitText(); Invalidate(); };
            _text.Enter += (s, e) => { _focusWithin = true; Invalidate(); };
            _text.MouseDown += TextMouseDown;
            _text.MouseMove += TextMouseMove;
            _text.MouseUp += TextMouseUp;
            _text.MouseEnter += (s, e) => { _hover = true; Invalidate(); };
            _text.MouseLeave += (s, e) => { if (!_scrubbing) { _hover = false; Invalidate(); } };
            _text.MouseWheel += SpinnerMouseWheel;
            _arrows.MouseWheel += SpinnerMouseWheel;

            Controls.Add(_text);
            Controls.Add(_arrows);

            MouseWheel += SpinnerMouseWheel;
            MouseEnter += (s, e) => { _hover = true; Invalidate(); };
            MouseLeave += (s, e) => { _hover = false; Invalidate(); };
            Resize += (s, e) => LayoutInner();

            Value = 0M;
        }

        public int DecimalPlaces
        {
            get { return _decimalPlaces; }
            set { _decimalPlaces = Math.Max(0, Math.Min(4, value)); RefreshText(); }
        }

        public decimal Minimum
        {
            get { return _minimum; }
            set { _minimum = value; if (_value < _minimum) Value = _minimum; }
        }

        public decimal Maximum
        {
            get { return _maximum; }
            set { _maximum = value; if (_value > _maximum) Value = _maximum; }
        }

        public decimal Increment
        {
            get { return _increment; }
            set { _increment = value <= 0 ? 1M : value; }
        }

        public decimal Value
        {
            get { return _value; }
            set
            {
                decimal v = Math.Max(_minimum, Math.Min(_maximum, value));
                if (_value == v)
                {
                    RefreshText();
                    return;
                }

                _value = v;
                RefreshText();
                if (ValueChanged != null) ValueChanged(this, EventArgs.Empty);
            }
        }

        public bool ThousandsSeparator { get; set; }

        public HorizontalAlignment TextAlign
        {
            get { return _text.TextAlign; }
            set { _text.TextAlign = value; }
        }

        private void LayoutInner()
        {
            _text.Left = 6;
            _text.Top = Math.Max(4, (Height - _text.PreferredHeight) / 2);
            _text.Width = Math.Max(30, Width - _arrows.Width - 11);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            Color border = (_focusWithin || _hover || _scrubbing)
                ? Color.FromArgb(192, 50, 45)
                : Color.FromArgb(214, 217, 222);

            using (var pen = new Pen(border, (_focusWithin || _scrubbing) ? 1.5f : 1f))
                e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
        }

        protected override void OnEnabledChanged(EventArgs e)
        {
            base.OnEnabledChanged(e);
            _text.Enabled = Enabled;
            _arrows.Enabled = Enabled;
            _text.BackColor = Enabled ? Color.White : Color.FromArgb(247, 247, 248);
            BackColor = _text.BackColor;
            Invalidate();
        }

        private void RefreshText()
        {
            if (_internalUpdate) return;
            _internalUpdate = true;
            try
            {
                _text.Text = _value.ToString("F" + _decimalPlaces, CultureInfo.CurrentCulture);
            }
            finally { _internalUpdate = false; }
        }

        private void CommitText()
        {
            if (_internalUpdate) return;
            decimal parsed;
            if (!decimal.TryParse(_text.Text, NumberStyles.Float, CultureInfo.CurrentCulture, out parsed))
            {
                string alt = (_text.Text ?? "").Replace(',', '.');
                if (!decimal.TryParse(alt, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed))
                {
                    RefreshText();
                    return;
                }
            }
            Value = parsed;
        }

        private void TextKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                CommitText();
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Up)
            {
                Step(1);
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Down)
            {
                Step(-1);
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                RefreshText();
                e.SuppressKeyPress = true;
            }
        }

        private void SpinnerMouseWheel(object sender, MouseEventArgs e)
        {
            if (!Enabled || e.Delta == 0) return;
            Step(e.Delta > 0 ? 1 : -1);
        }

        private void Step(int direction)
        {
            if (!Enabled) return;
            decimal mult = (ModifierKeys & Keys.Shift) == Keys.Shift ? 10M : 1M;
            Value += direction * _increment * mult;
        }

        private void TextMouseDown(object sender, MouseEventArgs e)
        {
            if (!Enabled || e.Button != MouseButtons.Left) return;
            _scrubbing = true;
            _scrubStartScreenY = Cursor.Position.Y;
            _scrubStartValue = _value;
            _text.Cursor = Cursors.SizeNS;
            _text.Capture = true;
            Invalidate();
        }

        private void TextMouseMove(object sender, MouseEventArgs e)
        {
            if (!_scrubbing || (e.Button & MouseButtons.Left) == 0) return;

            int dy = _scrubStartScreenY - Cursor.Position.Y;
            if (Math.Abs(dy) < 2) return;

            decimal mult = (ModifierKeys & Keys.Shift) == Keys.Shift ? 10M : 1M;
            decimal steps = Math.Round(dy / 4M, 0);
            Value = _scrubStartValue + steps * _increment * mult;
        }

        private void TextMouseUp(object sender, MouseEventArgs e)
        {
            if (!_scrubbing) return;
            _scrubbing = false;
            _text.Capture = false;
            _text.Cursor = Cursors.IBeam;
            Invalidate();
        }

        private sealed class SpinnerArrows : Control
        {
            public event EventHandler StepUp;
            public event EventHandler StepDown;
            private bool _hoverTop;
            private bool _hoverBottom;

            public SpinnerArrows()
            {
                DoubleBuffered = true;
                Cursor = Cursors.Hand;
                SetStyle(ControlStyles.ResizeRedraw | ControlStyles.UserPaint |
                         ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer, true);
            }

            protected override void OnMouseMove(MouseEventArgs e)
            {
                base.OnMouseMove(e);
                bool top = e.Y < Height / 2;
                bool bottom = !top;
                if (_hoverTop != top || _hoverBottom != bottom)
                {
                    _hoverTop = top;
                    _hoverBottom = bottom;
                    Invalidate();
                }
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                _hoverTop = _hoverBottom = false;
                Invalidate();
            }

            protected override void OnMouseDown(MouseEventArgs e)
            {
                base.OnMouseDown(e);
                if (!Enabled || e.Button != MouseButtons.Left) return;
                if (e.Y < Height / 2)
                {
                    if (StepUp != null) StepUp(this, EventArgs.Empty);
                }
                else
                {
                    if (StepDown != null) StepDown(this, EventArgs.Empty);
                }
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                base.OnPaint(e);
                var g = e.Graphics;
                g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

                if (_hoverTop)
                    using (var b = new SolidBrush(Color.FromArgb(252, 235, 234)))
                        g.FillRectangle(b, 0, 0, Width, Height / 2);
                if (_hoverBottom)
                    using (var b = new SolidBrush(Color.FromArgb(252, 235, 234)))
                        g.FillRectangle(b, 0, Height / 2, Width, Height - Height / 2);

                using (var p = new Pen(Color.FromArgb(226, 228, 232)))
                {
                    g.DrawLine(p, 0, 0, 0, Height);
                    g.DrawLine(p, 0, Height / 2, Width, Height / 2);
                }

                Color c = Enabled ? Color.FromArgb(102, 106, 112) : Color.FromArgb(180, 183, 188);
                using (var p = new Pen(c, 1.3f))
                {
                    int cx = Width / 2;
                    int ty = Height / 4;
                    g.DrawLine(p, cx - 3, ty + 1, cx, ty - 2);
                    g.DrawLine(p, cx, ty - 2, cx + 3, ty + 1);

                    int by = Height * 3 / 4;
                    g.DrawLine(p, cx - 3, by - 1, cx, by + 2);
                    g.DrawLine(p, cx, by + 2, cx + 3, by - 1);
                }
            }
        }
    }

    [ComVisible(true)]
    [Guid("D1F2AE42-A212-4A7B-8E03-3C9FE4F22B77")]
    [ProgId("OfficeSmartFontPicker.FontPickerPane")]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class FontPickerPane : UserControl
    {
        internal static FontPickerPane CurrentInstance;

        private List<FontFamilyInfo> _families;
        private List<FontFamilyInfo> _searchResults = new List<FontFamilyInfo>();
        private List<int> _weights = new List<int>();
        private FontFamilyInfo _selectedFamily;
        private FontFaceInfo _activeFace;

        private TableLayoutPanel _layout;
        private TextBox _search;
        private Panel _resultsPanel;
        private ListBox _results;
        private Button _fontDropButton;
        private WeightSlider _weight;
        private Label _weightValue;
        private Panel _formatStrip;
        private CheckBox _italicButton;
        private Button _caseButton;
        private CheckBox _subscriptButton;
        private CheckBox _superscriptButton;
        private ScrubSpinner _fontSizeInput;
        private ScrubSpinner _trackingInput;
        private ScrubSpinner _lineSpacingInput;
        private Label _lineSpacingUnit;
        private Label _status;
        private Label _liveBadge;
        private Button _fitTextBoxes;
        private Timer _selectionPollTimer;

        private bool _suspendLiveApply;
        private bool _syncControls;
        private bool _suppressSearch;
        private bool _lineSpacingUsesLines = true;
        private int _caseModeIndex = -1;
        private object _lastSelectionObject;
        private object _lastTextRange;
        private object _lastShapeRange;


        private static readonly Color CaseUpper = Color.FromArgb(192, 50, 45);
        private static readonly Color CaseSentence = Color.FromArgb(217, 126, 35);
        private static readonly Color CaseTitle = Color.FromArgb(54, 112, 191);

        private static readonly Color Bg = Color.FromArgb(244, 245, 247);
        private static readonly Color Card = Color.White;
        private static readonly Color TextMain = Color.FromArgb(38, 42, 48);
        private static readonly Color TextMuted = Color.FromArgb(112, 118, 128);
        private static readonly Color Accent = Color.FromArgb(192, 50, 45);
        private static readonly Color Border = Color.FromArgb(221, 224, 229);
        private static readonly Color LiveGreen = Color.FromArgb(192, 50, 45);

        public FontPickerPane()
        {
            CurrentInstance = this;
            Dock = DockStyle.Fill;
            BackColor = Bg;
            Font = new Font("Segoe UI", 9F);
            BuildUI();

            try
            {
                _status.Text = "正在读取本机字体…";
                Application.DoEvents();
                _families = FontCatalog.Load();
                ShowEmptySearchState();

                _selectionPollTimer = new Timer();
                _selectionPollTimer.Interval = 350;
                _selectionPollTimer.Tick += (s2, e) =>
                {
                    if (!ContainsFocus) SyncFromCurrentSelection();
                };
                _selectionPollTimer.Start();

                SyncFromCurrentSelection();
                Connect.Log("Pane V18 loaded, families=" + _families.Count);
            }
            catch (Exception ex)
            {
                _families = new List<FontFamilyInfo>();
                _status.Text = "字体读取失败";
                Connect.Log("Pane V18 load ERROR " + ex);
            }
        }

        protected override void Dispose(bool disposing)
        {
            try
            {
                if (_selectionPollTimer != null)
                {
                    _selectionPollTimer.Stop();
                    _selectionPollTimer.Dispose();
                    _selectionPollTimer = null;
                }
            }
            catch { }

            if (object.ReferenceEquals(CurrentInstance, this))
                CurrentInstance = null;
            base.Dispose(disposing);
        }

        internal void SyncFromCurrentSelection()
        {
            try
            {
                dynamic app = Connect.ApplicationInstance;
                if (app == null || app.ActiveWindow == null) return;
                SyncFromPowerPointSelection(app.ActiveWindow.Selection);
            }
            catch { }
        }

        internal void SyncFromPowerPointSelection(object selectionObject)
        {
            if (_families == null || selectionObject == null || IsDisposed) return;

            try
            {
                if (InvokeRequired)
                {
                    BeginInvoke(new Action<object>(SyncFromPowerPointSelection), selectionObject);
                    return;
                }
            }
            catch { }

            string fontName;
            bool bold;
            bool italic;
            bool mixed;
            if (!TryReadSelectionFont(selectionObject, out fontName, out bold, out italic, out mixed))
            {
                ShowNoTextSelection();
                return;
            }

            _lastSelectionObject = selectionObject;
            CacheSelectionTargets(selectionObject);

            if (mixed || string.IsNullOrWhiteSpace(fontName))
            {
                ShowMixedSelection();
                SyncInlineStylesFromSelection(selectionObject);
                SyncFontSizeFromSelection(selectionObject);
                SyncSpacingFromSelection(selectionObject);
                return;
            }

            FontFamilyInfo family;
            FontFaceInfo face;
            if (FindCatalogFace(fontName, bold, italic, out family, out face))
            {
                _selectedFamily = family;

                _syncControls = true;
                try
                {
                    _italicButton.Checked = italic;
                    _italicButton.Enabled = true;
                }
                finally { _syncControls = false; }

                _suppressSearch = true;
                _search.Text = family.Name;
                _search.SelectionStart = _search.Text.Length;
                _suppressSearch = false;
                SetResultsVisible(false, 0);

                SelectFaceExact(face, false);
                SyncInlineStylesFromSelection(selectionObject);
                SyncFontSizeFromSelection(selectionObject);
                SyncSpacingFromSelection(selectionObject);
                _status.Text = "";
            }
            else
            {
                _selectedFamily = null;
                _activeFace = null;

                _suppressSearch = true;
                _search.Text = fontName;
                _search.SelectionStart = _search.Text.Length;
                _suppressSearch = false;
                SetResultsVisible(false, 0);

                _weightValue.Text = bold ? "700  \u00b7  Bold" : "400  \u00b7  Regular";
                _weight.Enabled = false;

                _syncControls = true;
                try
                {
                    _italicButton.Checked = italic;
                    _italicButton.Enabled = true;
                }
                finally { _syncControls = false; }
                SyncInlineStylesFromSelection(selectionObject);
                SyncFontSizeFromSelection(selectionObject);
                SyncSpacingFromSelection(selectionObject);

                _status.Text = "\u5f53\u524d\u5b57\u4f53\u672a\u5efa\u7acb\u5b57\u91cd\u6620\u5c04";
            }
        }

        private bool TryReadSelectionFont(
            object selectionObject,
            out string fontName,
            out bool bold,
            out bool italic,
            out bool mixed)
        {
            fontName = null;
            bold = false;
            italic = false;
            mixed = false;

            try
            {
                dynamic sel = selectionObject;
                int type = (int)sel.Type;

                if (type == 3) // ppSelectionText
                {
                    dynamic font = sel.TextRange2.Font;
                    object nameObj = null;
                    try { nameObj = font.Name; } catch { }
                    fontName = nameObj == null ? null : Convert.ToString(nameObj);

                    int b = 0, it = 0;
                    try { b = (int)font.Bold; } catch { }
                    try { it = (int)font.Italic; } catch { }

                    // Office returns mixed tri-state values for mixed text.
                    mixed = b == -2 || it == -2 || string.IsNullOrWhiteSpace(fontName);
                    bold = b != 0 && b != -2;
                    italic = it != 0 && it != -2;
                    return true;
                }

                if (type == 2) // ppSelectionShapes
                {
                    dynamic range = sel.ShapeRange;
                    var names = new List<string>();
                    var bolds = new List<bool>();
                    var italics = new List<bool>();

                    for (int i = 1; i <= range.Count; i++)
                    {
                        dynamic shape = range[i];
                        try
                        {
                            if ((int)shape.HasTextFrame != -1) continue;
                            if ((int)shape.TextFrame.HasText != -1) continue;

                            dynamic font = shape.TextFrame2.TextRange.Font;
                            object nameObj = null;
                            try { nameObj = font.Name; } catch { }
                            string name = nameObj == null ? null : Convert.ToString(nameObj);
                            if (string.IsNullOrWhiteSpace(name))
                            {
                                mixed = true;
                                continue;
                            }

                            int b = 0, it = 0;
                            try { b = (int)font.Bold; } catch { }
                            try { it = (int)font.Italic; } catch { }
                            if (b == -2 || it == -2) mixed = true;

                            names.Add(name);
                            bolds.Add(b != 0 && b != -2);
                            italics.Add(it != 0 && it != -2);
                        }
                        catch { }
                    }

                    if (names.Count == 0) return false;

                    fontName = names[0];
                    bold = bolds.Count > 0 && bolds[0];
                    italic = italics.Count > 0 && italics[0];

                    string firstName = fontName;
                    bool firstBold = bold;
                    bool firstItalic = italic;

                    if (names.Any(n => !n.Equals(firstName, StringComparison.CurrentCultureIgnoreCase)))
                        mixed = true;
                    if (bolds.Any(x => x != firstBold) || italics.Any(x => x != firstItalic))
                        mixed = true;

                    return true;
                }
            }
            catch { }

            return false;
        }

        private bool FindCatalogFace(
            string fontName,
            bool bold,
            bool italic,
            out FontFamilyInfo family,
            out FontFaceInfo face)
        {
            family = null;
            face = null;
            if (string.IsNullOrWhiteSpace(fontName)) return false;

            string key = NormalizeSearch(fontName);

            foreach (var fam in _families)
            {
                var exactFace = fam.Faces.FirstOrDefault(f =>
                    NormalizeSearch(f.ApplyName).Equals(key, StringComparison.OrdinalIgnoreCase));
                if (exactFace != null)
                {
                    family = fam;
                    var usable = UsableFaces(fam);
                    var upright = usable.Where(f => !f.Italic).ToList();
                    if (upright.Count == 0) upright = usable;

                    face = upright
                        .OrderBy(f => Math.Abs(f.Weight - exactFace.Weight))
                        .ThenBy(f => VariantKey(f).Equals(VariantKey(exactFace), StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                        .ThenByDescending(f => f.ExactRegistered)
                        .FirstOrDefault() ?? exactFace;
                    return true;
                }
            }

            foreach (var fam in _families)
            {
                bool familyMatch =
                    NormalizeSearch(fam.Name).Equals(key, StringComparison.OrdinalIgnoreCase)
                    || fam.Aliases.Any(a =>
                        NormalizeSearch(a).Equals(key, StringComparison.OrdinalIgnoreCase));

                if (!familyMatch) continue;

                family = fam;
                var usable = UsableFaces(fam);
                int desiredWeight = bold ? 700 : 400;
                var styled = usable.Where(f => !f.Italic).ToList();
                if (styled.Count == 0) styled = usable;
                face = styled
                    .OrderBy(f => Math.Abs(f.Weight - desiredWeight))
                    .ThenBy(f => VariantKey(f).Length == 0 ? 0 : 1)
                    .ThenByDescending(f => f.ExactRegistered)
                    .FirstOrDefault();
                return face != null;
            }

            return false;
        }

        private void ShowNoTextSelection()
        {
            _selectedFamily = null;
            _activeFace = null;
            _weightValue.Text = "\u2014";
            _weight.Enabled = false;

            if (_italicButton != null) _italicButton.Enabled = false;
            if (_caseButton != null) _caseButton.Enabled = false;
            if (_subscriptButton != null) _subscriptButton.Enabled = false;
            if (_superscriptButton != null) _superscriptButton.Enabled = false;
            if (_fontSizeInput != null) _fontSizeInput.Enabled = false;
            if (_trackingInput != null) _trackingInput.Enabled = false;
            if (_lineSpacingInput != null) _lineSpacingInput.Enabled = false;

            _status.Text = "\u7b49\u5f85\u6587\u5b57\u9009\u62e9";
        }

        private void ShowMixedSelection()
        {
            _selectedFamily = null;
            _activeFace = null;
            _suppressSearch = true;
            _search.Text = "\u591a\u79cd\u5b57\u4f53";
            _suppressSearch = false;
            SetResultsVisible(false, 0);

            _weightValue.Text = "\u2014";
            _weight.Enabled = false;

            if (_italicButton != null) _italicButton.Enabled = true;
            if (_caseButton != null) _caseButton.Enabled = true;
            if (_subscriptButton != null) _subscriptButton.Enabled = true;
            if (_superscriptButton != null) _superscriptButton.Enabled = true;
            if (_fontSizeInput != null) _fontSizeInput.Enabled = true;
            if (_trackingInput != null) _trackingInput.Enabled = true;
            if (_lineSpacingInput != null) _lineSpacingInput.Enabled = true;

            _status.Text = "\u5f53\u524d\u9009\u533a\u5305\u542b\u591a\u79cd\u5b57\u4f53\u6216\u5b57\u91cd";
        }

        private void BuildUI()
        {
            _layout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 3,
                BackColor = Card,
                Padding = new Padding(0)
            };
            _layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 78));
            _layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 0));
            _layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));

            var header = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Card,
                Padding = new Padding(12, 7, 12, 7)
            };

            var topRow = new Panel
            {
                Dock = DockStyle.Top,
                Height = 22,
                BackColor = Bg
            };

            var title = new Label
            {
                Text = "\u5b57\u4f53",
                Dock = DockStyle.Left,
                Width = 140,
                Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                ForeColor = TextMain,
                TextAlign = ContentAlignment.MiddleLeft
            };

            _liveBadge = new Label
            {
                Text = "DESIGN BY SUKIN",
                Dock = DockStyle.Right,
                Width = 118,
                Font = new Font("Segoe UI Semibold", 7.5F, FontStyle.Bold),
                ForeColor = LiveGreen,
                TextAlign = ContentAlignment.MiddleRight
            };

            topRow.Controls.Add(_liveBadge);
            topRow.Controls.Add(title);

            var searchShell = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 36,
                BackColor = Card,
                Padding = new Padding(10, 7, 0, 5)
            };
            searchShell.Paint += (s2, e) =>
            {
                Color borderColor = _search != null && _search.Focused ? Accent : Border;
                using (var pen = new Pen(borderColor, _search != null && _search.Focused ? 1.5f : 1f))
                    e.Graphics.DrawRectangle(pen, 0, 0, searchShell.Width - 1, searchShell.Height - 1);
            };

            _fontDropButton = new Button
            {
                Dock = DockStyle.Right,
                Width = 36,
                Text = "",
                FlatStyle = FlatStyle.Flat,
                BackColor = Card,
                ForeColor = TextMuted,
                TabStop = false,
                Cursor = Cursors.Hand
            };
            _fontDropButton.FlatAppearance.BorderSize = 0;
            _fontDropButton.Paint += DrawDropdownChevron;
            _fontDropButton.Click += (s2, e) => ToggleFontDropdown();

            _search = new TextBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.None,
                Font = new Font("Segoe UI", 10F),
                BackColor = Card,
                ForeColor = TextMain
            };
            _search.TextChanged += (s2, e) =>
            {
                if (!_suppressSearch) SearchFonts(_search.Text);
            };
            _search.KeyDown += SearchKeyDown;
            _search.GotFocus += (s2, e) =>
            {
                searchShell.Invalidate();
                if (_selectedFamily != null && _search.Text == _selectedFamily.Name)
                    _search.SelectAll();
            };
            _search.LostFocus += (s2, e) => searchShell.Invalidate();

            searchShell.Controls.Add(_search);
            searchShell.Controls.Add(_fontDropButton);

            header.Controls.Add(searchShell);
            header.Controls.Add(topRow);

            _resultsPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Bg,
                Padding = new Padding(12, 0, 12, 8)
            };

            _results = new ListBox
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                IntegralHeight = false,
                DrawMode = DrawMode.OwnerDrawFixed,
                ItemHeight = 38,
                BackColor = Card,
                ForeColor = TextMain
            };
            _results.DrawItem += DrawResultItem;
            _results.Click += (s2, e) => SelectSearchResult();
            _results.DoubleClick += (s2, e) => SelectSearchResult();
            _results.KeyDown += ResultsKeyDown;
            _resultsPanel.Controls.Add(_results);

            var editor = BuildEditor();

            _layout.Controls.Add(header, 0, 0);
            _layout.Controls.Add(_resultsPanel, 0, 1);
            _layout.Controls.Add(editor, 0, 2);
            Controls.Add(_layout);
        }

        private Control BuildEditor()
        {
            var panel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = Card,
                AutoScroll = true
            };

            // WEIGHT
            var weightSection = new Panel
            {
                Left = 0,
                Top = 0,
                Width = 360,
                Height = 104,
                BackColor = Card
            };

            var weightTitle = new Label
            {
                Text = "WEIGHT  \u5b57\u91cd",
                Left = 16,
                Top = 14,
                Width = 120,
                Height = 20,
                Font = new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold),
                ForeColor = TextMuted
            };
            weightSection.Controls.Add(weightTitle);

            _weightValue = new Label
            {
                Text = "\u2014",
                Left = 184,
                Top = 11,
                Width = 156,
                Height = 24,
                TextAlign = ContentAlignment.MiddleRight,
                Font = new Font("Segoe UI Semibold", 9.5F, FontStyle.Bold),
                ForeColor = TextMain
            };
            weightSection.Controls.Add(_weightValue);

            _weight = new WeightSlider
            {
                Left = 12,
                Top = 38,
                Width = 336,
                Height = 60,
                Minimum = 0,
                Maximum = 0,
                AccentColor = Accent,
                TrackColor = Color.FromArgb(218, 221, 225),
                LabelColor = TextMuted,
                Enabled = false
            };
            _weight.ValueChanged += (s2, e) =>
            {
                if (!_suspendLiveApply && !_syncControls) SliderChanged();
            };
            weightSection.Controls.Add(_weight);

            panel.Controls.Add(weightSection);
            panel.Controls.Add(MakeSeparator(103));

            // FONT SIZE
            var sizeSection = new Panel
            {
                Left = 0,
                Top = 104,
                Width = 360,
                Height = 54,
                BackColor = Card
            };

            var sizeTitle = new Label
            {
                Text = "SIZE  \u5b57\u53f7",
                Left = 16,
                Top = 15,
                Width = 120,
                Height = 20,
                Font = new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold),
                ForeColor = TextMuted
            };
            sizeSection.Controls.Add(sizeTitle);

            _fontSizeInput = new ScrubSpinner
            {
                Left = 190,
                Top = 12,
                Width = 116,
                Height = 30,
                DecimalPlaces = 1,
                Minimum = 1.0M,
                Maximum = 400.0M,
                Increment = 0.5M,
                Value = 18.0M,
                Enabled = false
            };
            _fontSizeInput.ValueChanged += (s2, e) =>
            {
                if (!_syncControls && _fontSizeInput.Enabled)
                    ApplyFontSizeFromSpinner();
            };
            sizeSection.Controls.Add(_fontSizeInput);

            var sizeUnit = new Label
            {
                Text = "pt",
                Left = 312,
                Top = 12,
                Width = 30,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = TextMuted
            };
            sizeSection.Controls.Add(sizeUnit);

            panel.Controls.Add(sizeSection);
            panel.Controls.Add(MakeSeparator(157));

            // FORMAT TOOLS
            var formatSection = new Panel
            {
                Left = 0,
                Top = 158,
                Width = 360,
                Height = 92,
                BackColor = Card
            };
            _formatStrip = formatSection;

            int cellWidth = 86;

            _italicButton = MakeToolToggle("I", 8, "\u659c\u4f53");
            _italicButton.Font = new Font("Segoe UI", 13F, FontStyle.Italic);
            _italicButton.CheckedChanged += (s2, e) =>
            {
                UpdateToggleVisual(_italicButton);
                if (!_syncControls) ItalicChanged();
            };
            formatSection.Controls.Add(_italicButton);
            formatSection.Controls.Add(MakeToolCaption("\u659c\u4f53", 8));

            _caseButton = MakeToolButton("Aa", 94, "\u5927\u5c0f\u5199");
            _caseButton.Click += (s2, e) => CycleCaseMode();
            _caseButton.Paint += DrawCaseModeIndicator;
            formatSection.Controls.Add(_caseButton);
            formatSection.Controls.Add(MakeToolCaption("\u5927\u5c0f\u5199", 94));

            _subscriptButton = MakeToolToggle("", 180, "\u4e0b\u6807");
            _subscriptButton.Paint += (s2, e) => DrawBaselineIcon(_subscriptButton, e.Graphics, false);
            _subscriptButton.CheckedChanged += (s2, e) =>
            {
                UpdateToggleVisual(_subscriptButton);
                if (_syncControls) return;
                if (_subscriptButton.Checked)
                {
                    _syncControls = true;
                    try
                    {
                        _superscriptButton.Checked = false;
                        UpdateToggleVisual(_superscriptButton);
                    }
                    finally { _syncControls = false; }
                }
                ApplyBaselineStyle(true, _subscriptButton.Checked);
            };
            formatSection.Controls.Add(_subscriptButton);
            formatSection.Controls.Add(MakeToolCaption("\u4e0b\u6807", 180));

            _superscriptButton = MakeToolToggle("", 266, "\u4e0a\u6807");
            _superscriptButton.Paint += (s2, e) => DrawBaselineIcon(_superscriptButton, e.Graphics, true);
            _superscriptButton.CheckedChanged += (s2, e) =>
            {
                UpdateToggleVisual(_superscriptButton);
                if (_syncControls) return;
                if (_superscriptButton.Checked)
                {
                    _syncControls = true;
                    try
                    {
                        _subscriptButton.Checked = false;
                        UpdateToggleVisual(_subscriptButton);
                    }
                    finally { _syncControls = false; }
                }
                ApplyBaselineStyle(false, _superscriptButton.Checked);
            };
            formatSection.Controls.Add(_superscriptButton);
            formatSection.Controls.Add(MakeToolCaption("\u4e0a\u6807", 266));

            for (int i = 1; i < 4; i++)
            {
                var divider = new Panel
                {
                    Left = i * cellWidth,
                    Top = 15,
                    Width = 1,
                    Height = 58,
                    BackColor = Color.FromArgb(236, 237, 239)
                };
                formatSection.Controls.Add(divider);
            }

            panel.Controls.Add(formatSection);
            panel.Controls.Add(MakeSeparator(249));

            // SPACING
            var spacingSection = new Panel
            {
                Left = 0,
                Top = 250,
                Width = 360,
                Height = 128,
                BackColor = Card
            };

            var spacingTitle = new Label
            {
                Text = "SPACING  \u95f4\u8ddd",
                Left = 16,
                Top = 13,
                Width = 135,
                Height = 20,
                Font = new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold),
                ForeColor = TextMuted
            };
            spacingSection.Controls.Add(spacingTitle);

            var trackingLabel = new Label
            {
                Text = "\u5b57\u95f4\u8ddd",
                Left = 16,
                Top = 44,
                Width = 90,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei UI", 9F),
                ForeColor = TextMain
            };
            spacingSection.Controls.Add(trackingLabel);

            _trackingInput = new ScrubSpinner
            {
                Left = 190,
                Top = 44,
                Width = 116,
                Height = 30,
                DecimalPlaces = 1,
                Minimum = -20M,
                Maximum = 100M,
                Increment = 0.1M,
                Value = 0M,
                Enabled = false
            };
            _trackingInput.ValueChanged += (s2, e) =>
            {
                if (!_syncControls && _trackingInput.Enabled)
                    ApplyTrackingFromSpinner();
            };
            spacingSection.Controls.Add(_trackingInput);

            var trackingUnit = new Label
            {
                Text = "pt",
                Left = 312,
                Top = 44,
                Width = 30,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = TextMuted
            };
            spacingSection.Controls.Add(trackingUnit);

            var innerSep = new Panel
            {
                Left = 16,
                Top = 83,
                Width = 328,
                Height = 1,
                BackColor = Color.FromArgb(240, 240, 242)
            };
            spacingSection.Controls.Add(innerSep);

            var lineLabel = new Label
            {
                Text = "\u884c\u95f4\u8ddd",
                Left = 16,
                Top = 92,
                Width = 90,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                Font = new Font("Microsoft YaHei UI", 9F),
                ForeColor = TextMain
            };
            spacingSection.Controls.Add(lineLabel);

            _lineSpacingInput = new ScrubSpinner
            {
                Left = 190,
                Top = 92,
                Width = 116,
                Height = 30,
                DecimalPlaces = 2,
                Minimum = 0.10M,
                Maximum = 10.00M,
                Increment = 0.05M,
                Value = 1.00M,
                Enabled = false
            };
            _lineSpacingInput.ValueChanged += (s2, e) =>
            {
                if (!_syncControls && _lineSpacingInput.Enabled)
                    ApplyLineSpacingFromSpinner();
            };
            spacingSection.Controls.Add(_lineSpacingInput);

            _lineSpacingUnit = new Label
            {
                Text = "\u00d7",
                Left = 312,
                Top = 92,
                Width = 30,
                Height = 30,
                TextAlign = ContentAlignment.MiddleLeft,
                ForeColor = TextMuted
            };
            spacingSection.Controls.Add(_lineSpacingUnit);

            _status = new Label
            {
                Visible = false,
                Width = 1,
                Height = 1
            };

            panel.Controls.Add(spacingSection);
            panel.Controls.Add(MakeSeparator(377));

            // SLIDE TOOLS
            var toolsSection = new Panel
            {
                Left = 0,
                Top = 378,
                Width = 360,
                Height = 96,
                BackColor = Card
            };

            var toolTitle = new Label
            {
                Text = "SLIDE TOOLS  \u9875\u9762\u5de5\u5177",
                Left = 16,
                Top = 13,
                Width = 180,
                Height = 20,
                Font = new Font("Segoe UI Semibold", 8.5F, FontStyle.Bold),
                ForeColor = TextMuted
            };
            toolsSection.Controls.Add(toolTitle);

            _fitTextBoxes = new Button
            {
                Text = "\u672c\u9875\u6587\u5b57\u6846\u8d34\u5408\u6587\u5b57",
                Left = 16,
                Top = 43,
                Width = 328,
                Height = 36,
                FlatStyle = FlatStyle.Flat,
                BackColor = Color.White,
                ForeColor = Accent,
                Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Cursor = Cursors.Hand
            };
            _fitTextBoxes.FlatAppearance.BorderColor = Color.FromArgb(227, 154, 151);
            _fitTextBoxes.FlatAppearance.MouseOverBackColor = Color.FromArgb(253, 241, 240);
            _fitTextBoxes.FlatAppearance.MouseDownBackColor = Color.FromArgb(248, 224, 222);
            _fitTextBoxes.Click += (s2, e) => FitCurrentSlideTextBoxes();
            toolsSection.Controls.Add(_fitTextBoxes);

            panel.Controls.Add(toolsSection);

            panel.Resize += (s2, e) =>
            {
                int w = Math.Max(270, panel.ClientSize.Width);

                weightSection.Width = w;
                sizeSection.Width = w;
                formatSection.Width = w;
                spacingSection.Width = w;
                toolsSection.Width = w;

                _weightValue.Left = Math.Max(154, w - 176);
                _weightValue.Width = Math.Max(120, w - _weightValue.Left - 16);
                _weight.Width = Math.Max(220, w - 24);

                int inputLeft = Math.Max(138, w - 170);
                _fontSizeInput.Left = inputLeft;
                sizeUnit.Left = Math.Min(w - 34, inputLeft + 122);

                int usable = Math.Max(240, w - 16);
                int cell = usable / 4;
                PositionToolCell(_italicButton, 8, cell);
                PositionToolCell(_caseButton, 8 + cell, cell);
                PositionToolCell(_subscriptButton, 8 + cell * 2, cell);
                PositionToolCell(_superscriptButton, 8 + cell * 3, cell);

                _trackingInput.Left = inputLeft;
                _lineSpacingInput.Left = inputLeft;
                trackingUnit.Left = Math.Min(w - 34, inputLeft + 122);
                _lineSpacingUnit.Left = Math.Min(w - 34, inputLeft + 122);
                innerSep.Width = Math.Max(180, w - 32);

                _fitTextBoxes.Width = Math.Max(180, w - 32);
            };

            return panel;
        }

        private Panel MakeSeparator(int top)
        {
            return new Panel
            {
                Left = 16,
                Top = top,
                Width = 328,
                Height = 1,
                BackColor = Color.FromArgb(232, 233, 236),
                Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right
            };
        }

        private CheckBox MakeToolToggle(string text, int left, string tooltip)
        {
            var control = new CheckBox
            {
                Text = text,
                Appearance = Appearance.Button,
                Left = left,
                Top = 10,
                Width = 60,
                Height = 40,
                FlatStyle = FlatStyle.Flat,
                BackColor = Card,
                ForeColor = TextMain,
                TextAlign = ContentAlignment.MiddleCenter,
                AutoCheck = true,
                TabStop = false,
                Enabled = false,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI", 11F)
            };
            control.FlatAppearance.BorderSize = 0;
            control.FlatAppearance.MouseOverBackColor = Color.FromArgb(248, 248, 249);
            control.FlatAppearance.MouseDownBackColor = Color.FromArgb(247, 221, 219);
            control.FlatAppearance.CheckedBackColor = Color.FromArgb(253, 238, 237);

            var tip = new ToolTip();
            tip.SetToolTip(control, tooltip);
            return control;
        }

        private Button MakeToolButton(string text, int left, string tooltip)
        {
            var control = new Button
            {
                Text = text,
                Left = left,
                Top = 10,
                Width = 60,
                Height = 40,
                FlatStyle = FlatStyle.Flat,
                BackColor = Card,
                ForeColor = TextMain,
                TextAlign = ContentAlignment.MiddleCenter,
                TabStop = false,
                Enabled = false,
                Cursor = Cursors.Hand,
                Font = new Font("Segoe UI Semibold", 10.5F)
            };
            control.FlatAppearance.BorderSize = 0;
            control.FlatAppearance.MouseOverBackColor = Color.FromArgb(248, 248, 249);
            control.FlatAppearance.MouseDownBackColor = Color.FromArgb(247, 221, 219);

            var tip = new ToolTip();
            tip.SetToolTip(control, tooltip);
            return control;
        }

        private Label MakeToolCaption(string text, int left)
        {
            return new Label
            {
                Text = text,
                Left = left,
                Top = 57,
                Width = 68,
                Height = 20,
                TextAlign = ContentAlignment.MiddleCenter,
                Font = new Font("Microsoft YaHei UI", 8F),
                ForeColor = TextMuted
            };
        }

        private void PositionToolCell(Control control, int left, int cellWidth)
        {
            if (control == null) return;
            control.Left = left + Math.Max(0, (cellWidth - 60) / 2);
            control.Width = Math.Min(66, Math.Max(52, cellWidth - 8));
        }

        private void DrawBaselineIcon(CheckBox control, Graphics g, bool superscript)
        {
            if (control == null) return;

            Color color = control.Checked ? Accent : TextMain;
            using (var mainFont = new Font("Segoe UI", 12F, FontStyle.Regular))
            using (var smallFont = new Font("Segoe UI", 7F, FontStyle.Regular))
            {
                string digit = superscript ? "2" : "1";
                Size mainSize = TextRenderer.MeasureText("T", mainFont, Size.Empty, TextFormatFlags.NoPadding);
                Size smallSize = TextRenderer.MeasureText(digit, smallFont, Size.Empty, TextFormatFlags.NoPadding);

                int totalW = mainSize.Width + smallSize.Width - 2;
                int x = (control.ClientSize.Width - totalW) / 2;
                int mainY = (control.ClientSize.Height - mainSize.Height) / 2 + 1;
                int smallY = superscript ? mainY - 3 : mainY + mainSize.Height - smallSize.Height + 2;

                TextRenderer.DrawText(g, "T", mainFont, new Point(x, mainY), color, TextFormatFlags.NoPadding);
                TextRenderer.DrawText(g, digit, smallFont, new Point(x + mainSize.Width - 2, smallY), color, TextFormatFlags.NoPadding);
            }
        }

        private void UpdateToggleVisual(CheckBox button)
        {
            if (button == null) return;
            button.ForeColor = button.Checked ? Accent : TextMain;
            button.BackColor = button.Checked ? Color.FromArgb(253, 238, 237) : Card;
            button.Invalidate();
        }

        private void CycleCaseMode()
        {
            _caseModeIndex = (_caseModeIndex + 1) % 3;
            int officeMode;

            if (_caseModeIndex == 0) officeMode = 3;       // UPPERCASE
            else if (_caseModeIndex == 1) officeMode = 1;  // Sentence case
            else officeMode = 4;                           // Title Case

            UpdateCaseButtonVisual();
            ChangeCaseSelection(officeMode);
        }

        private void UpdateCaseButtonVisual()
        {
            if (_caseButton == null) return;

            Color accent;
            string mode;

            if (_caseModeIndex == 0)
            {
                accent = CaseUpper;
                mode = "\u5168\u90e8\u5927\u5199";
            }
            else if (_caseModeIndex == 1)
            {
                accent = CaseSentence;
                mode = "\u53e5\u9996\u5b57\u6bcd\u5927\u5199";
            }
            else if (_caseModeIndex == 2)
            {
                accent = CaseTitle;
                mode = "\u6bcf\u4e2a\u5355\u8bcd\u9996\u5b57\u6bcd\u5927\u5199";
            }
            else
            {
                accent = TextMain;
                mode = "\u5927\u5c0f\u5199\u5faa\u73af";
            }

            _caseButton.ForeColor = accent;
            _caseButton.BackColor = Card;
            _caseButton.Invalidate();

            var tip = new ToolTip();
            tip.SetToolTip(_caseButton, "\u5f53\u524d\uff1a" + mode + "\uff1b\u70b9\u51fb\u5207\u6362\u4e0b\u4e00\u79cd");
        }

        private void DrawCaseModeIndicator(object sender, PaintEventArgs e)
        {
            if (_caseModeIndex < 0) return;

            Color accent = _caseModeIndex == 0
                ? CaseUpper
                : (_caseModeIndex == 1 ? CaseSentence : CaseTitle);

            var c = sender as Control;
            if (c == null) return;

            using (var b = new SolidBrush(accent))
                e.Graphics.FillRectangle(b, 12, c.Height - 3, Math.Max(8, c.Width - 24), 2);
        }

        private void ApplyTrackingFromSpinner()
        {
            int count = ApplyTrackingToSelection((float)_trackingInput.Value);
            if (count > 0)
                _status.Text = "\u5b57\u95f4\u8ddd  " + _trackingInput.Value.ToString("0.0") + " pt";
        }

        private void ApplyLineSpacingFromSpinner()
        {
            float value = (float)_lineSpacingInput.Value;
            int count = ApplyLineSpacingToSelection(value, _lineSpacingUsesLines);
            if (count > 0)
            {
                _status.Text = _lineSpacingUsesLines
                    ? "\u884c\u95f4\u8ddd  " + value.ToString("0.00") + "\u00d7"
                    : "\u884c\u95f4\u8ddd  " + value.ToString("0.0") + " pt";
            }
        }

        private void DrawDropdownChevron(object sender, PaintEventArgs e)
        {
            var button = sender as Button;
            if (button == null) return;

            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            int cx = button.ClientRectangle.Width / 2;
            int cy = button.ClientRectangle.Height / 2;
            using (var pen = new Pen(TextMuted, 1.6f))
            {
                e.Graphics.DrawLine(pen, cx - 4, cy - 2, cx, cy + 2);
                e.Graphics.DrawLine(pen, cx, cy + 2, cx + 4, cy - 2);
            }
        }

        private void DrawCardBorder(object sender, PaintEventArgs e)
        {
            var c = sender as Control;
            if (c == null) return;
            using (var pen = new Pen(Border))
                e.Graphics.DrawRectangle(pen, 0, 0, c.Width - 1, c.Height - 1);
        }

        private void DrawResultItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0 || e.Index >= _searchResults.Count) return;

            var fam = _searchResults[e.Index];
            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            var bg = selected ? Color.FromArgb(253, 238, 237) : Card;
            var fg = selected ? Accent : TextMain;

            using (var b = new SolidBrush(bg))
                e.Graphics.FillRectangle(b, e.Bounds);

            var nameRect = new Rectangle(e.Bounds.Left + 10, e.Bounds.Top + 3, Math.Max(100, e.Bounds.Width - 72), 32);
            TextRenderer.DrawText(
                e.Graphics,
                fam.Name,
                new Font("Segoe UI", 9.5F, selected ? FontStyle.Bold : FontStyle.Regular),
                nameRect,
                fg,
                TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            try
            {
                using (var sampleFont = new Font(fam.Name, 13F, FontStyle.Regular, GraphicsUnit.Point))
                {
                    var sampleRect = new Rectangle(e.Bounds.Right - 58, e.Bounds.Top + 3, 46, 32);
                    TextRenderer.DrawText(
                        e.Graphics,
                        "Aa",
                        sampleFont,
                        sampleRect,
                        fg,
                        TextFormatFlags.VerticalCenter | TextFormatFlags.HorizontalCenter);
                }
            }
            catch { }

            using (var pen = new Pen(Color.FromArgb(238, 238, 238)))
                e.Graphics.DrawLine(pen, e.Bounds.Left, e.Bounds.Bottom - 1, e.Bounds.Right, e.Bounds.Bottom - 1);
        }

        private static string NormalizeSearch(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "";
            var chars = s.ToLowerInvariant()
                .Where(ch => char.IsLetterOrDigit(ch))
                .ToArray();
            return new string(chars);
        }

        private static int SubsequencePenalty(string text, string query)
        {
            int qi = 0, gaps = 0, last = -1;
            for (int i = 0; i < text.Length && qi < query.Length; i++)
            {
                if (text[i] == query[qi])
                {
                    if (last >= 0) gaps += i - last - 1;
                    last = i;
                    qi++;
                }
            }
            return qi == query.Length ? gaps : int.MaxValue;
        }

        private static int SearchScore(FontFamilyInfo f, string rawQuery)
        {
            string q = NormalizeSearch(rawQuery);
            if (q.Length == 0) return int.MaxValue;

            var names = new List<string> { f.Name };
            names.AddRange(f.Aliases);

            int best = int.MaxValue;
            foreach (var raw in names)
            {
                string n = NormalizeSearch(raw);
                if (n.Length == 0) continue;

                if (n.StartsWith(q)) best = Math.Min(best, 0);
                else
                {
                    int pos = n.IndexOf(q, StringComparison.Ordinal);
                    if (pos >= 0) best = Math.Min(best, 10 + pos);
                    else
                    {
                        int gap = SubsequencePenalty(n, q);
                        if (gap != int.MaxValue)
                            best = Math.Min(best, 40 + gap);
                    }
                }
            }
            return best;
        }

        private void SearchFonts(string query)
        {
            if (_families == null) return;
            query = (query ?? "").Trim();

            if (query.Length == 0)
            {
                ShowEmptySearchState();
                return;
            }

            _searchResults = _families
                .Select(f => new { Font = f, Score = SearchScore(f, query) })
                .Where(x => x.Score != int.MaxValue)
                .OrderBy(x => x.Score)
                .ThenBy(x => x.Font.Name, StringComparer.CurrentCultureIgnoreCase)
                .Take(30)
                .Select(x => x.Font)
                .ToList();

            _results.BeginUpdate();
            _results.Items.Clear();
            foreach (var f in _searchResults)
                _results.Items.Add(f.Name);
            _results.EndUpdate();

            int resultHeight = Math.Min(210, Math.Max(54, _searchResults.Count * 40 + 2));
            SetResultsVisible(true, resultHeight);

            _status.Text = _searchResults.Count == 0
                ? "没有匹配的字体"
                : "找到 " + _searchResults.Count + " 个字体家族";
        }

        private void ToggleFontDropdown()
        {
            if (_families == null || _families.Count == 0) return;

            if (_layout.RowStyles[1].Height > 0)
            {
                SetResultsVisible(false, 0);
                return;
            }

            _searchResults = _families
                .OrderBy(f => f.Name, StringComparer.CurrentCultureIgnoreCase)
                .ToList();

            _results.BeginUpdate();
            _results.Items.Clear();
            foreach (var f in _searchResults)
                _results.Items.Add(f.Name);
            _results.EndUpdate();

            SetResultsVisible(true, 230);
            _status.Text = "\u5168\u90e8 " + _searchResults.Count + " \u4e2a\u5b57\u4f53\u5bb6\u65cf";
            _results.Focus();

            if (_selectedFamily != null)
            {
                int index = _searchResults.FindIndex(
                    f => f.Name.Equals(_selectedFamily.Name, StringComparison.CurrentCultureIgnoreCase));
                if (index >= 0)
                {
                    _results.SelectedIndex = index;
                    _results.TopIndex = Math.Max(0, index - 2);
                }
            }
        }

        private void ShowEmptySearchState()
        {
            _searchResults.Clear();
            _results.Items.Clear();
            SetResultsVisible(false, 0);

            if (_selectedFamily == null)
                _status.Text = "\u9009\u4e2d\u6587\u5b57\u540e\u53c2\u6570\u5b9e\u65f6\u540c\u6b65";
            else
                _status.Text = "";
        }

        private void SetResultsVisible(bool visible, int height)
        {
            _layout.RowStyles[1].Height = visible ? height : 0;
            _resultsPanel.Visible = visible;
        }

        private void SearchKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down && _results.Items.Count > 0)
            {
                _results.Focus();
                _results.SelectedIndex = Math.Max(0, _results.SelectedIndex);
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Enter && _results.Items.Count > 0)
            {
                if (_results.SelectedIndex < 0) _results.SelectedIndex = 0;
                SelectSearchResult();
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                _search.Clear();
                e.SuppressKeyPress = true;
            }
        }

        private void ResultsKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                SelectSearchResult();
                e.SuppressKeyPress = true;
            }
            else if (e.KeyCode == Keys.Escape)
            {
                _search.Focus();
                SetResultsVisible(false, 0);
                e.SuppressKeyPress = true;
            }
        }

        private void SelectSearchResult()
        {
            int i = _results.SelectedIndex;
            if (i < 0 || i >= _searchResults.Count) return;
            SelectFamily(_searchResults[i]);
        }

        private List<FontFaceInfo> UsableFaces(FontFamilyInfo family)
        {
            if (family == null) return new List<FontFaceInfo>();

            var usable = family.Faces
                .Where(f => f.ExactRegistered)
                .ToList();

            // Variable fonts often expose many named instances through DirectWrite,
            // while PowerPoint COM can only address the family plus Bold/Italic reliably.
            // Keep only basic Regular/Bold faces for those non-registered instances.
            foreach (var face in family.Faces.Where(f => !f.ExactRegistered))
            {
                bool basicWeight = face.Weight == 400 || face.Weight == 700;
                bool basicVariant = VariantKey(face).Length == 0;
                if (basicWeight && basicVariant && !usable.Contains(face))
                    usable.Add(face);
            }

            if (usable.Count == 0)
            {
                var basic = family.Faces
                    .Where(f => (f.Weight == 400 || f.Weight == 700) && VariantKey(f).Length == 0)
                    .ToList();
                if (basic.Count > 0) usable.AddRange(basic);
                else usable.AddRange(family.Faces.Take(1));
            }

            return usable;
        }

        private void SelectFamily(FontFamilyInfo family)
        {
            _selectedFamily = family;

            var usableFaces = UsableFaces(family);

            _syncControls = true;
            try
            {
                _italicButton.Enabled = true;

                _suppressSearch = true;
                _search.Text = family.Name;
                _search.SelectionStart = _search.Text.Length;
                _suppressSearch = false;
                SetResultsVisible(false, 0);
            }
            finally
            {
                _syncControls = false;
            }

            var defaultFace = usableFaces
                .Where(f => !f.Italic)
                .OrderBy(f => Math.Abs(f.Weight - 400))
                .ThenBy(f => VariantKey(f).Length == 0 ? 0 : 1)
                .ThenBy(f => f.DisplayFace.Equals("Regular", StringComparison.OrdinalIgnoreCase) ? 0 : 1)
                .ThenByDescending(f => f.ExactRegistered)
                .ThenBy(f => f.DisplayFace, StringComparer.CurrentCultureIgnoreCase)
                .FirstOrDefault()
                ?? usableFaces
                    .OrderBy(f => Math.Abs(f.Weight - 400))
                    .ThenBy(f => VariantKey(f).Length == 0 ? 0 : 1)
                    .ThenByDescending(f => f.ExactRegistered)
                    .FirstOrDefault();

            if (defaultFace != null)
                SelectFaceExact(defaultFace, true);

            _search.Focus();
        }

        private List<FontFaceInfo> CurrentStyleFaces()
        {
            if (_selectedFamily == null) return new List<FontFaceInfo>();
            var usable = UsableFaces(_selectedFamily);
            var upright = usable.Where(f => !f.Italic).ToList();
            return upright.Count > 0 ? upright : usable;
        }

        private void RebuildWeightChoices(int desiredWeight)
        {
            var styleFaces = CurrentStyleFaces();
            _weights = styleFaces
                .Select(f => f.Weight)
                .Distinct()
                .OrderBy(w => w)
                .ToList();

            if (_weights.Count == 0) _weights.Add(desiredWeight > 0 ? desiredWeight : 400);

            _weight.Minimum = 0;
            _weight.Maximum = Math.Max(0, _weights.Count - 1);
            _weight.TickFrequency = 1;
            _weight.SmallChange = 1;
            _weight.LargeChange = 1;
            _weight.SetLabels(_weights);
            _weight.Enabled = _weights.Count > 1;

            int bestIndex = 0;
            int bestDistance = int.MaxValue;
            for (int i = 0; i < _weights.Count; i++)
            {
                int d = Math.Abs(_weights[i] - desiredWeight);
                if (d < bestDistance)
                {
                    bestDistance = d;
                    bestIndex = i;
                }
            }
            _weight.Value = Math.Min(bestIndex, _weight.Maximum);
        }

        private void SelectFaceExact(FontFaceInfo face, bool liveApply)
        {
            if (face == null) return;
            _activeFace = face;

            _syncControls = true;
            _suspendLiveApply = true;
            try
            {
                RebuildWeightChoices(face.Weight);
            }
            finally
            {
                _suspendLiveApply = false;
                _syncControls = false;
            }

            RefreshActiveFace(liveApply);
        }

        private static string VariantKey(FontFaceInfo face)
        {
            if (face == null) return "";
            var weightTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "regular","normal","roman","book",
                "thin","hairline","extralight","ultralight","light",
                "medium","semibold","demibold","demi","bold",
                "extrabold","ultrabold","black","heavy",
                "italic","oblique",
                "el","l","r","m","sb","b","eb","h"
            };

            var tokens = (face.DisplayFace ?? "")
                .Replace("-", " ")
                .Replace("_", " ")
                .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Where(t => !weightTokens.Contains(t))
                .Select(t => t.ToLowerInvariant());

            return string.Join(" ", tokens);
        }

        private void SliderChanged()
        {
            if (_selectedFamily == null || _weights.Count == 0) return;

            int weight = _weights[Math.Max(0, Math.Min(_weight.Value, _weights.Count - 1))];
            var candidates = CurrentStyleFaces().Where(f => f.Weight == weight).ToList();
            if (candidates.Count == 0) return;

            string variant = VariantKey(_activeFace);
            var sameVariant = candidates
                .Where(f => VariantKey(f).Equals(variant, StringComparison.OrdinalIgnoreCase))
                .ToList();
            if (sameVariant.Count > 0) candidates = sameVariant;

            var face = candidates
                .OrderByDescending(f => f.ExactRegistered)
                .ThenBy(f => f.DisplayFace, StringComparer.CurrentCultureIgnoreCase)
                .First();

            _activeFace = face;

            RefreshActiveFace(true);
        }

        private void CacheSelectionTargets(object selectionObject)
        {
            if (selectionObject == null) return;
            try
            {
                dynamic sel = selectionObject;
                int type = (int)sel.Type;

                if (type == 3)
                {
                    _lastTextRange = sel.TextRange2;
                    _lastShapeRange = null;
                }
                else if (type == 2)
                {
                    _lastShapeRange = sel.ShapeRange;
                    _lastTextRange = null;
                }
            }
            catch { }
        }

        private void GetWorkingTargets(out object textRange, out object shapeRange)
        {
            textRange = null;
            shapeRange = null;

            try
            {
                dynamic app = Connect.ApplicationInstance;
                if (app != null && app.ActiveWindow != null)
                {
                    dynamic current = app.ActiveWindow.Selection;
                    int type = (int)current.Type;

                    if (type == 3)
                    {
                        textRange = current.TextRange2;
                        return;
                    }
                    if (type == 2)
                    {
                        shapeRange = current.ShapeRange;
                        return;
                    }
                }
            }
            catch { }

            textRange = _lastTextRange;
            shapeRange = _lastShapeRange;
        }

        private void ItalicChanged()
        {
            if (_syncControls) return;
            ApplyItalicNative(_italicButton.Checked);
        }

        private void ApplyItalicNative(bool enabled)
        {
            try
            {
                object textObj, shapeObj;
                GetWorkingTargets(out textObj, out shapeObj);
                int count = 0;

                if (textObj != null)
                {
                    dynamic textRange = textObj;
                    try { textRange.Font.Italic = enabled ? -1 : 0; count = 1; } catch { }
                }
                else if (shapeObj != null)
                {
                    dynamic range = shapeObj;
                    for (int i = 1; i <= range.Count; i++)
                        count += ApplyItalicToShape(range[i], enabled);
                }

                _status.Text = count > 0
                    ? (enabled ? "\u5df2\u542f\u7528\u659c\u4f53" : "\u5df2\u53d6\u6d88\u659c\u4f53")
                    : "\u5f53\u524d\u9009\u533a\u6ca1\u6709\u6587\u5b57";
            }
            catch (Exception ex)
            {
                Connect.Log("Italic native ERROR " + ex);
            }
        }

        private static int ApplyItalicToShape(dynamic shape, bool enabled)
        {
            int count = 0;
            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        count += ApplyItalicToShape(items.Item(i), enabled);
                    return count;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                            {
                                cellShape.TextFrame2.TextRange.Font.Italic = enabled ? -1 : 0;
                                count++;
                            }
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                {
                    shape.TextFrame2.TextRange.Font.Italic = enabled ? -1 : 0;
                    count++;
                }
            }
            catch { }

            return count;
        }

        private void RefreshActiveFace(bool liveApply)
        {
            if (_activeFace == null) return;

            _weightValue.Text = _activeFace.Weight + "  \u00B7  " + _activeFace.DisplayFace;

            if (liveApply && !_suspendLiveApply)
                TryApplyLive();
        }

        private void ChangeCaseSelection(int caseMode)
        {
            try
            {
                object textObj, shapeObj;
                GetWorkingTargets(out textObj, out shapeObj);
                int count = 0;

                if (textObj != null)
                {
                    dynamic textRange = textObj;
                    textRange.ChangeCase((Office.MsoTextChangeCase)caseMode);
                    count = 1;
                }
                else if (shapeObj != null)
                {
                    dynamic range = shapeObj;
                    for (int i = 1; i <= range.Count; i++)
                        count += ChangeCaseInShape(range[i], caseMode);
                }

                _status.Text = count > 0
                    ? "\u5df2\u66f4\u6539\u5927\u5c0f\u5199"
                    : "\u5f53\u524d\u9009\u533a\u6ca1\u6709\u6587\u5b57";
            }
            catch (Exception ex)
            {
                Connect.Log("ChangeCase ERROR " + ex);
            }
        }

        private static int ChangeCaseInShape(dynamic shape, int caseMode)
        {
            int count = 0;
            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        count += ChangeCaseInShape(items.Item(i), caseMode);
                    return count;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                            {
                                cellShape.TextFrame2.TextRange.ChangeCase((Office.MsoTextChangeCase)caseMode);
                                count++;
                            }
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                {
                    shape.TextFrame2.TextRange.ChangeCase((Office.MsoTextChangeCase)caseMode);
                    count++;
                }
            }
            catch { }

            return count;
        }

        private void ApplyBaselineStyle(bool subscript, bool enabled)
        {
            try
            {
                object textObj, shapeObj;
                GetWorkingTargets(out textObj, out shapeObj);
                int count = 0;

                if (textObj != null)
                {
                    dynamic textRange = textObj;
                    ApplyBaselineToTextRange(textRange, subscript, enabled);
                    count = 1;
                }
                else if (shapeObj != null)
                {
                    dynamic range = shapeObj;
                    for (int i = 1; i <= range.Count; i++)
                        count += ApplyBaselineToShape(range[i], subscript, enabled);
                }

                _status.Text = count > 0
                    ? (subscript
                        ? (enabled ? "\u5df2\u542f\u7528\u4e0b\u6807" : "\u5df2\u53d6\u6d88\u4e0b\u6807")
                        : (enabled ? "\u5df2\u542f\u7528\u4e0a\u6807" : "\u5df2\u53d6\u6d88\u4e0a\u6807"))
                    : "\u5f53\u524d\u9009\u533a\u6ca1\u6709\u6587\u5b57";
            }
            catch (Exception ex)
            {
                Connect.Log("Baseline style ERROR " + ex);
            }
        }

        private static void ApplyBaselineToTextRange(dynamic textRange, bool subscript, bool enabled)
        {
            try
            {
                dynamic font = textRange.Font;

                if (!enabled)
                {
                    try { font.Subscript = 0; } catch { }
                    try { font.Superscript = 0; } catch { }
                    try { font.BaselineOffset = 0.0f; } catch { }
                    return;
                }

                if (subscript)
                {
                    try { font.Superscript = 0; } catch { }
                    try { font.Subscript = -1; } catch { }
                    try { font.BaselineOffset = -0.25f; } catch { }
                }
                else
                {
                    try { font.Subscript = 0; } catch { }
                    try { font.Superscript = -1; } catch { }
                    try { font.BaselineOffset = 0.30f; } catch { }
                }
            }
            catch { }
        }

        private static int ApplyBaselineToShape(dynamic shape, bool subscript, bool enabled)
        {
            int count = 0;

            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        count += ApplyBaselineToShape(items.Item(i), subscript, enabled);
                    return count;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                            {
                                ApplyBaselineToTextRange(cellShape.TextFrame2.TextRange, subscript, enabled);
                                count++;
                            }
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                {
                    ApplyBaselineToTextRange(shape.TextFrame2.TextRange, subscript, enabled);
                    count++;
                }
            }
            catch { }

            return count;
        }

        private void SyncInlineStylesFromSelection(object selectionObject)
        {
            if (_italicButton == null || selectionObject == null) return;

            var italics = new List<int>();
            var subs = new List<int>();
            var supers = new List<int>();

            try
            {
                dynamic sel = selectionObject;
                int type = (int)sel.Type;

                if (type == 3)
                {
                    CollectInlineStyleFromTextRange(sel.TextRange2, italics, subs, supers);
                }
                else if (type == 2)
                {
                    dynamic range = sel.ShapeRange;
                    for (int i = 1; i <= range.Count; i++)
                        CollectInlineStyleFromShape(range[i], italics, subs, supers);
                }
            }
            catch { }

            bool hasText = italics.Count > 0 || subs.Count > 0 || supers.Count > 0;

            _syncControls = true;
            try
            {
                _italicButton.Enabled = hasText;
                _caseButton.Enabled = hasText;
                _subscriptButton.Enabled = hasText;
                _superscriptButton.Enabled = hasText;

                if (hasText)
                {
                    _italicButton.Checked = italics.Count > 0 && italics.All(x => x != 0 && x != -2);
                    _subscriptButton.Checked = subs.Count > 0 && subs.All(x => x != 0 && x != -2);
                    _superscriptButton.Checked = supers.Count > 0 && supers.All(x => x != 0 && x != -2);
                }
                else
                {
                    _italicButton.Checked = false;
                    _subscriptButton.Checked = false;
                    _superscriptButton.Checked = false;
                }

                UpdateToggleVisual(_italicButton);
                UpdateToggleVisual(_subscriptButton);
                UpdateToggleVisual(_superscriptButton);
            }
            finally
            {
                _syncControls = false;
            }
        }

        private static void CollectInlineStyleFromTextRange(
            dynamic textRange,
            List<int> italics,
            List<int> subs,
            List<int> supers)
        {
            try { italics.Add((int)textRange.Font.Italic); } catch { }

            int sub = 0;
            int sup = 0;
            bool gotSub = false;
            bool gotSup = false;

            try { sub = (int)textRange.Font.Subscript; gotSub = true; } catch { }
            try { sup = (int)textRange.Font.Superscript; gotSup = true; } catch { }

            try
            {
                float offset = (float)textRange.Font.BaselineOffset;
                if (offset < -0.01f)
                {
                    sub = -1;
                    sup = 0;
                    gotSub = gotSup = true;
                }
                else if (offset > 0.01f)
                {
                    sup = -1;
                    sub = 0;
                    gotSub = gotSup = true;
                }
            }
            catch { }

            if (gotSub) subs.Add(sub);
            if (gotSup) supers.Add(sup);
        }

        private static void CollectInlineStyleFromShape(
            dynamic shape,
            List<int> italics,
            List<int> subs,
            List<int> supers)
        {
            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        CollectInlineStyleFromShape(items.Item(i), italics, subs, supers);
                    return;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                                CollectInlineStyleFromTextRange(
                                    cellShape.TextFrame2.TextRange,
                                    italics,
                                    subs,
                                    supers);
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                    CollectInlineStyleFromTextRange(
                        shape.TextFrame2.TextRange,
                        italics,
                        subs,
                        supers);
            }
            catch { }
        }

        private void SyncFontSizeFromSelection(object selectionObject)
        {
            if (_fontSizeInput == null || selectionObject == null) return;

            var sizes = new List<float>();

            try
            {
                dynamic sel = selectionObject;
                int type = (int)sel.Type;

                if (type == 3)
                {
                    CollectFontSizeFromTextRange(sel.TextRange2, sizes);
                }
                else if (type == 2)
                {
                    dynamic range = sel.ShapeRange;
                    for (int i = 1; i <= range.Count; i++)
                        CollectFontSizeFromShape(range[i], sizes);
                }
            }
            catch { }

            _syncControls = true;
            try
            {
                bool hasText = sizes.Count > 0;
                _fontSizeInput.Enabled = hasText;
                if (!hasText) return;

                float first = sizes[0];
                bool same = sizes.All(x => Math.Abs(x - first) < 0.01f);
                if (same)
                    SetNumericValueSafe(_fontSizeInput, (decimal)first);
            }
            finally
            {
                _syncControls = false;
            }
        }

        private static void CollectFontSizeFromTextRange(dynamic textRange, List<float> sizes)
        {
            try
            {
                float size = (float)textRange.Font.Size;
                if (!float.IsNaN(size) && !float.IsInfinity(size) && size > 0f && size < 10000f)
                    sizes.Add(size);
            }
            catch { }
        }

        private static void CollectFontSizeFromShape(dynamic shape, List<float> sizes)
        {
            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        CollectFontSizeFromShape(items.Item(i), sizes);
                    return;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                                CollectFontSizeFromTextRange(cellShape.TextFrame2.TextRange, sizes);
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                    CollectFontSizeFromTextRange(shape.TextFrame2.TextRange, sizes);
            }
            catch { }
        }

        private void ApplyFontSizeFromSpinner()
        {
            float value = (float)_fontSizeInput.Value;
            int count = ApplyFontSizeToSelection(value);
            if (count > 0)
                _status.Text = "\u5b57\u53f7  " + value.ToString("0.0") + " pt";
        }

        private int ApplyFontSizeToSelection(float value)
        {
            try
            {
                object textObj, shapeObj;
                GetWorkingTargets(out textObj, out shapeObj);

                if (textObj != null)
                {
                    dynamic textRange = textObj;
                    textRange.Font.Size = value;
                    return 1;
                }

                if (shapeObj != null)
                {
                    int count = 0;
                    dynamic range = shapeObj;
                    for (int i = 1; i <= range.Count; i++)
                        count += ApplyFontSizeToShape(range[i], value);
                    return count;
                }
            }
            catch (Exception ex)
            {
                Connect.Log("Font size apply ERROR " + ex);
            }

            return 0;
        }

        private static int ApplyFontSizeToShape(dynamic shape, float value)
        {
            int count = 0;

            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        count += ApplyFontSizeToShape(items.Item(i), value);
                    return count;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                            {
                                cellShape.TextFrame2.TextRange.Font.Size = value;
                                count++;
                            }
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                {
                    shape.TextFrame2.TextRange.Font.Size = value;
                    count++;
                }
            }
            catch { }

            return count;
        }

        private void SyncSpacingFromSelection(object selectionObject)
        {
            if (_trackingInput == null || _lineSpacingInput == null || selectionObject == null)
                return;

            var trackings = new List<float>();
            var lineValues = new List<float>();
            var lineRules = new List<int>();

            try
            {
                dynamic sel = selectionObject;
                int type = (int)sel.Type;

                if (type == 3)
                {
                    CollectSpacingFromTextRange(sel.TextRange2, trackings, lineValues, lineRules);
                }
                else if (type == 2)
                {
                    dynamic range = sel.ShapeRange;
                    for (int i = 1; i <= range.Count; i++)
                        CollectSpacingFromShape(range[i], trackings, lineValues, lineRules);
                }
            }
            catch { }

            _syncControls = true;
            try
            {
                bool hasText = trackings.Count > 0 || lineValues.Count > 0;
                _trackingInput.Enabled = hasText;
                _lineSpacingInput.Enabled = hasText;

                if (!hasText)
                    return;

                if (trackings.Count > 0)
                {
                    float first = trackings[0];
                    bool same = trackings.All(x => Math.Abs(x - first) < 0.01f);
                    if (same)
                        SetNumericValueSafe(_trackingInput, (decimal)first);
                }

                if (lineValues.Count > 0 && lineRules.Count > 0)
                {
                    float firstValue = lineValues[0];
                    int firstRule = lineRules[0];
                    bool sameValue = lineValues.All(x => Math.Abs(x - firstValue) < 0.01f);
                    bool sameRule = lineRules.All(x => x == firstRule);

                    if (sameValue && sameRule && firstRule != -2)
                    {
                        _lineSpacingUsesLines = firstRule != 0;
                        ConfigureLineSpacingControl(_lineSpacingUsesLines);
                        SetNumericValueSafe(_lineSpacingInput, (decimal)firstValue);
                        _lineSpacingUnit.Text = _lineSpacingUsesLines ? "\u00d7" : "pt";
                    }
                }
            }
            finally
            {
                _syncControls = false;
            }
        }

        private static void SetNumericValueSafe(ScrubSpinner control, decimal value)
        {
            if (control == null) return;
            if (value < control.Minimum) value = control.Minimum;
            if (value > control.Maximum) value = control.Maximum;
            control.Value = value;
        }

        private void ConfigureLineSpacingControl(bool useLines)
        {
            if (_lineSpacingInput == null) return;

            decimal current = _lineSpacingInput.Value;

            if (useLines)
            {
                _lineSpacingInput.DecimalPlaces = 2;
                _lineSpacingInput.Minimum = 0.10M;
                _lineSpacingInput.Maximum = 10.00M;
                _lineSpacingInput.Increment = 0.05M;
            }
            else
            {
                _lineSpacingInput.DecimalPlaces = 1;
                _lineSpacingInput.Minimum = 1.0M;
                _lineSpacingInput.Maximum = 500.0M;
                _lineSpacingInput.Increment = 0.5M;
            }

            SetNumericValueSafe(_lineSpacingInput, current);
        }

        private static void CollectSpacingFromTextRange(
            dynamic textRange,
            List<float> trackings,
            List<float> lineValues,
            List<int> lineRules)
        {
            try
            {
                float spacing = (float)textRange.Font.Spacing;
                if (!float.IsNaN(spacing) && !float.IsInfinity(spacing) && Math.Abs(spacing) < 10000f)
                    trackings.Add(spacing);
            }
            catch { }

            try
            {
                dynamic paragraph = textRange.ParagraphFormat;
                float spaceWithin = (float)paragraph.SpaceWithin;
                int lineRule = (int)paragraph.LineRuleWithin;

                if (!float.IsNaN(spaceWithin) && !float.IsInfinity(spaceWithin) && Math.Abs(spaceWithin) < 10000f)
                {
                    lineValues.Add(spaceWithin);
                    lineRules.Add(lineRule);
                }
            }
            catch { }
        }

        private static void CollectSpacingFromShape(
            dynamic shape,
            List<float> trackings,
            List<float> lineValues,
            List<int> lineRules)
        {
            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        CollectSpacingFromShape(items.Item(i), trackings, lineValues, lineRules);
                    return;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                                CollectSpacingFromTextRange(
                                    cellShape.TextFrame2.TextRange,
                                    trackings,
                                    lineValues,
                                    lineRules);
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                    CollectSpacingFromTextRange(
                        shape.TextFrame2.TextRange,
                        trackings,
                        lineValues,
                        lineRules);
            }
            catch { }
        }

        private int ApplyTrackingToSelection(float value)
        {
            try
            {
                object textObj, shapeObj;
                GetWorkingTargets(out textObj, out shapeObj);

                if (textObj != null)
                {
                    dynamic textRange = textObj;
                    textRange.Font.Spacing = value;
                    return 1;
                }

                if (shapeObj != null)
                {
                    int count = 0;
                    dynamic range = shapeObj;
                    for (int i = 1; i <= range.Count; i++)
                        count += ApplyTrackingToShape(range[i], value);
                    return count;
                }
            }
            catch (Exception ex)
            {
                Connect.Log("Tracking apply ERROR " + ex);
            }

            return 0;
        }

        private int ApplyLineSpacingToSelection(float value, bool useLines)
        {
            try
            {
                object textObj, shapeObj;
                GetWorkingTargets(out textObj, out shapeObj);

                if (textObj != null)
                {
                    dynamic textRange = textObj;
                    dynamic p = textRange.ParagraphFormat;
                    p.LineRuleWithin = useLines ? -1 : 0;
                    p.SpaceWithin = value;
                    return 1;
                }

                if (shapeObj != null)
                {
                    int count = 0;
                    dynamic range = shapeObj;
                    for (int i = 1; i <= range.Count; i++)
                        count += ApplyLineSpacingToShape(range[i], value, useLines);
                    return count;
                }
            }
            catch (Exception ex)
            {
                Connect.Log("Line spacing apply ERROR " + ex);
            }

            return 0;
        }

        private static int ApplyTrackingToShape(dynamic shape, float value)
        {
            int count = 0;

            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        count += ApplyTrackingToShape(items.Item(i), value);
                    return count;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                            {
                                cellShape.TextFrame2.TextRange.Font.Spacing = value;
                                count++;
                            }
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                {
                    shape.TextFrame2.TextRange.Font.Spacing = value;
                    count++;
                }
            }
            catch { }

            return count;
        }

        private static int ApplyLineSpacingToShape(dynamic shape, float value, bool useLines)
        {
            int count = 0;

            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        count += ApplyLineSpacingToShape(items.Item(i), value, useLines);
                    return count;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                            {
                                dynamic p = cellShape.TextFrame2.TextRange.ParagraphFormat;
                                p.LineRuleWithin = useLines ? -1 : 0;
                                p.SpaceWithin = value;
                                count++;
                            }
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                {
                    dynamic p = shape.TextFrame2.TextRange.ParagraphFormat;
                    p.LineRuleWithin = useLines ? -1 : 0;
                    p.SpaceWithin = value;
                    count++;
                }
            }
            catch { }

            return count;
        }

        private static string WeightName(int weight)
        {
            if (weight <= 150) return "Thin";
            if (weight <= 250) return "ExtraLight";
            if (weight <= 350) return "Light";
            if (weight <= 450) return "Regular";
            if (weight <= 550) return "Medium";
            if (weight <= 650) return "SemiBold";
            if (weight <= 750) return "Bold";
            if (weight <= 850) return "ExtraBold";
            return "Black";
        }

        private void FitCurrentSlideTextBoxes()
        {
            try
            {
                int count = SlideTextBoxTools.FitCurrentSlide(Connect.ApplicationInstance);
                _status.Text = count > 0
                    ? "\u5df2\u8d34\u5408 " + count + " \u4e2a\u6587\u672c\u6846"
                    : "\u672c\u9875\u6ca1\u6709\u53ef\u8c03\u6574\u7684\u666e\u901a\u6587\u672c\u6846";
            }
            catch (Exception ex)
            {
                _status.Text = "\u6587\u672c\u6846\u8c03\u6574\u5931\u8d25";
                Connect.Log("FitCurrentSlideTextBoxes V18 ERROR " + ex);
            }
        }

        private void TryApplyLive()
        {
            var face = _activeFace;
            if (face == null) return;

            try
            {
                object textObj, shapeObj;
                GetWorkingTargets(out textObj, out shapeObj);
                int count = 0;

                if (textObj != null)
                {
                    dynamic textRange = textObj;
                    ApplyFont(textRange.Font, face, _italicButton.Checked);
                    count = 1;
                }
                else if (shapeObj != null)
                {
                    dynamic range = shapeObj;
                    for (int i = 1; i <= range.Count; i++)
                        count += ApplyToShape(range[i], face, _italicButton.Checked);
                }
                else
                {
                    _status.Text = "\u8bf7\u5148\u9009\u4e2d\u6587\u5b57\u6216\u6587\u5b57\u6846";
                    return;
                }

                _status.Text = count > 0
                    ? face.Family + "  \u00b7  " + face.DisplayFace
                    : "\u5f53\u524d\u9009\u533a\u6ca1\u6709\u6587\u5b57";
            }
            catch (Exception ex)
            {
                _status.Text = "\u6682\u65f6\u65e0\u6cd5\u5e94\u7528";
                Connect.Log("LiveApply V18 ERROR " + ex);
            }
        }

        private static void ApplyFont(dynamic font, FontFaceInfo face, bool italicWanted)
        {
            try
            {
                if (face.ExactRegistered)
                {
                    try { font.Bold = 0; } catch { }
                    font.Name = face.ApplyName;
                    try { font.Italic = italicWanted ? -1 : 0; } catch { }
                    return;
                }

                font.Name = face.Family;
                try { font.Bold = face.Weight >= 700 ? -1 : 0; } catch { }
                try { font.Italic = italicWanted ? -1 : 0; } catch { }
            }
            catch { }
        }

        private static int ApplyToShape(dynamic shape, FontFaceInfo face, bool italicWanted)
        {
            int count = 0;

            try
            {
                if ((int)shape.Type == 6)
                {
                    dynamic items = shape.GroupItems;
                    for (int i = 1; i <= items.Count; i++)
                        count += ApplyToShape(items.Item(i), face, italicWanted);
                    return count;
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTable == -1)
                {
                    dynamic table = shape.Table;
                    for (int r = 1; r <= table.Rows.Count; r++)
                        for (int c = 1; c <= table.Columns.Count; c++)
                        {
                            dynamic cellShape = table.Cell(r, c).Shape;
                            if ((int)cellShape.TextFrame.HasText == -1)
                            {
                                ApplyFont(cellShape.TextFrame2.TextRange.Font, face, italicWanted);
                                count++;
                            }
                        }
                }
            }
            catch { }

            try
            {
                if ((int)shape.HasTextFrame == -1 && (int)shape.TextFrame.HasText == -1)
                {
                    ApplyFont(shape.TextFrame2.TextRange.Font, face, italicWanted);
                    count++;
                }
            }
            catch { }

            return count;
        }
    }
}
