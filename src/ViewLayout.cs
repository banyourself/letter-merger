using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

namespace LetterMerger
{
    public static class ViewLayout
    {
        public static TableLayoutPanel Rows()
        {
            var table = new TableLayoutPanel { ColumnCount = 1, AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, Dock = DockStyle.Top, Margin = Padding.Empty };
            table.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));
            return table;
        }

        public static void Add(TableLayoutPanel table, Control control, bool grow)
        {
            int row = table.RowCount++;
            table.RowStyles.Add(new RowStyle(grow ? SizeType.Percent : SizeType.AutoSize, grow ? 100 : 0));
            control.Dock = DockStyle.Fill;
            control.Margin = new Padding(0, 0, 0, 10);
            table.Controls.Add(control, 0, row);
        }

        public static Label Text(string text)
        {
            return new Label { Text = text, AutoSize = true, ForeColor = MainForm.Ink, Dock = DockStyle.Fill, Margin = new Padding(0, 6, 8, 6) };
        }

        public static FlowLayoutPanel Flow(params Control[] controls)
        {
            var panel = new FlowLayoutPanel { AutoSize = true, AutoSizeMode = AutoSizeMode.GrowAndShrink, WrapContents = true, Dock = DockStyle.Fill, Margin = Padding.Empty };
            foreach (var control in controls)
            {
                control.Dock = DockStyle.None;
                control.Margin = new Padding(0, 3, 10, 3);
                panel.Controls.Add(control);
            }
            return panel;
        }
    }

    public sealed class PreviewViewState
    {
        public const double MinimumZoom = .10, MaximumZoom = 8;
        public Size ImageSize { get; private set; }
        public Size Viewport { get; private set; }
        public bool FitMode { get; private set; }
        public double ManualZoom { get; private set; }
        public double PanX { get; private set; }
        public double PanY { get; private set; }

        public PreviewViewState()
        {
            FitMode = true;
            ManualZoom = 1;
            Viewport = new Size(1, 1);
        }

        public double Zoom
        {
            get
            {
                if (!FitMode || ImageSize.Width <= 0 || ImageSize.Height <= 0) return ManualZoom;
                return Math.Min(MaximumZoom, Math.Min(Viewport.Width / (double)ImageSize.Width, Viewport.Height / (double)ImageSize.Height));
            }
        }

        public RectangleF Bounds
        {
            get
            {
                double width = ImageSize.Width * Zoom, height = ImageSize.Height * Zoom;
                return new RectangleF((float)((Viewport.Width - width) / 2 + PanX), (float)((Viewport.Height - height) / 2 + PanY), (float)width, (float)height);
            }
        }

        public void SetImageSize(Size size)
        {
            if (size.Width < 0 || size.Height < 0) throw new ArgumentOutOfRangeException("size");
            ImageSize = size;
            PanX = PanY = 0;
            ClampPan();
        }

        public void SetViewport(Size size)
        {
            Viewport = new Size(Math.Max(1, size.Width), Math.Max(1, size.Height));
            ClampPan();
        }

        public void Fit()
        {
            FitMode = true;
            PanX = PanY = 0;
        }

        public void SetZoom(double requested)
        {
            SetZoomAt(requested, new PointF(Viewport.Width / 2f, Viewport.Height / 2f));
        }

        public void SetZoomAt(double requested, PointF pointer)
        {
            if (Double.IsNaN(requested) || Double.IsInfinity(requested) || requested <= 0 || Single.IsNaN(pointer.X) || Single.IsNaN(pointer.Y) || Single.IsInfinity(pointer.X) || Single.IsInfinity(pointer.Y)) throw new ArgumentOutOfRangeException("requested");
            double oldZoom = Zoom;
            var old = Bounds;
            double imageX = (pointer.X - old.X) / oldZoom, imageY = (pointer.Y - old.Y) / oldZoom;
            FitMode = false;
            ManualZoom = Math.Max(MinimumZoom, Math.Min(MaximumZoom, requested));
            PanX = pointer.X - Viewport.Width / 2.0 - (imageX - ImageSize.Width / 2.0) * Zoom;
            PanY = pointer.Y - Viewport.Height / 2.0 - (imageY - ImageSize.Height / 2.0) * Zoom;
            ClampPan();
        }

        public void Pan(double x, double y)
        {
            if (Double.IsNaN(x) || Double.IsInfinity(x) || Double.IsNaN(y) || Double.IsInfinity(y)) throw new ArgumentOutOfRangeException("x");
            if (FitMode) return;
            PanX += x;
            PanY += y;
            ClampPan();
        }

        void ClampPan()
        {
            double x = Math.Max(0, (ImageSize.Width * Zoom - Viewport.Width) / 2), y = Math.Max(0, (ImageSize.Height * Zoom - Viewport.Height) / 2);
            PanX = Math.Max(-x, Math.Min(x, PanX));
            PanY = Math.Max(-y, Math.Min(y, PanY));
            if (FitMode) PanX = PanY = 0;
        }
    }

    public sealed class ZoomPhotoView : Control
    {
        readonly PreviewViewState state = new PreviewViewState();
        Image image;
        Point dragStart;
        bool dragging;
        public event EventHandler ViewChanged;
        public PreviewViewState State { get { return state; } }

        public ZoomPhotoView()
        {
            SetStyle(ControlStyles.AllPaintingInWmPaint | ControlStyles.OptimizedDoubleBuffer | ControlStyles.UserPaint | ControlStyles.ResizeRedraw | ControlStyles.Selectable, true);
            TabStop = true;
            BackColor = ColorTranslator.FromHtml("#E5E7EB");
            MinimumSize = new Size(200, 140);
            AccessibleName = "Photo preview; mouse wheel to zoom, drag to pan";
        }

        public Image Image
        {
            get { return image; }
            set
            {
                if (ReferenceEquals(image, value)) return;
                var old = image;
                image = value;
                state.SetImageSize(image == null ? Size.Empty : image.Size);
                if (old != null) old.Dispose();
                Changed();
            }
        }

        public void Fit()
        {
            state.Fit();
            Changed();
        }

        public void SetZoom(double zoom)
        {
            state.SetZoom(zoom);
            Changed();
        }

        public void ZoomBy(double factor)
        {
            SetZoom(state.Zoom * factor);
        }

        void Changed()
        {
            Invalidate();
            if (ViewChanged != null) ViewChanged(this, EventArgs.Empty);
        }

        protected override void OnResize(EventArgs e)
        {
            base.OnResize(e);
            if (state == null || ClientSize.Width <= 24 || ClientSize.Height <= 24) return;
            state.SetViewport(new Size(ClientSize.Width - 24, ClientSize.Height - 24));
            Changed();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            if (image == null)
            {
                TextRenderer.DrawText(e.Graphics, "No photo to display for this row.", Font, ClientRectangle, MainForm.Ink, TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.WordBreak);
                return;
            }
            var bounds = state.Bounds;
            bounds.Offset(12, 12);
            e.Graphics.InterpolationMode = InterpolationMode.HighQualityBicubic;
            e.Graphics.PixelOffsetMode = PixelOffsetMode.HighQuality;
            e.Graphics.FillRectangle(Brushes.White, bounds);
            e.Graphics.DrawImage(image, bounds);
        }

        protected override void OnMouseEnter(EventArgs e)
        {
            base.OnMouseEnter(e);
            Focus();
        }

        protected override void OnMouseWheel(MouseEventArgs e)
        {
            base.OnMouseWheel(e);
            if (image == null) return;
            state.SetZoomAt(state.Zoom * Math.Pow(1.2, e.Delta / 120.0), new PointF(e.X - 12, e.Y - 12));
            Changed();
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            Focus();
            if (e.Button != MouseButtons.Left || image == null) return;
            dragging = true;
            dragStart = e.Location;
            Capture = true;
            Cursor = Cursors.SizeAll;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);
            if (!dragging) return;
            state.Pan(e.X - dragStart.X, e.Y - dragStart.Y);
            dragStart = e.Location;
            Changed();
        }

        protected override void OnMouseUp(MouseEventArgs e)
        {
            base.OnMouseUp(e);
            if (e.Button != MouseButtons.Left) return;
            Capture = false;
            dragging = false;
            Cursor = Cursors.Default;
        }

        protected override void OnMouseCaptureChanged(EventArgs e)
        {
            base.OnMouseCaptureChanged(e);
            if (Capture) return;
            dragging = false;
            Cursor = Cursors.Default;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing && image != null)
            {
                image.Dispose();
                image = null;
            }
            base.Dispose(disposing);
        }
    }
}
