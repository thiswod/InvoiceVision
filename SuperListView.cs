using ListView = System.Windows.Forms.ListView;

namespace InvoiceVision
{
    public class SuperListView : ListView
    {
        public SuperListView()
        {
            // 优化双缓冲设置
            this.DoubleBuffered = true;
            SetStyle(ControlStyles.OptimizedDoubleBuffer |
                     ControlStyles.AllPaintingInWmPaint, true);
            // 简化OwnerDraw设置
            this.OwnerDraw = true;
            this.DrawItem += (s, e) => e.DrawDefault = true;
            this.DrawSubItem += (s, e) => e.DrawDefault = true;
            this.DrawColumnHeader += (s, e) => e.DrawDefault = true;
            this.View = View.Details;//固定为详情列表视图
            //显示网格线
            this.GridLines = true;
            //整行选择 
            this.FullRowSelect = true;
        }
    }
}