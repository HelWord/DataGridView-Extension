using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using System.Drawing;
using System.ComponentModel;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using DevComponents.SuperGrid;
using DevComponents.DotNetBar.SuperGrid;
namespace GeneralAutoTestUI
{
    public static class SuperGridControlExtension
    {
        public static void GenerateColumnFromClass(this SuperGridControl ctl, Type target)
        {
            /*****************************反射生成各列*******************************************/
            /* 根据类的属性/属性/属性生成SuperGridControl的列                                          */
            /* 列名为属性的Description Attribute                                                                 */
            /* 是否显示为列根据属性的Browsable Attribute                                                   */
            /* 根据类的属性/属性/属性生成SuperGridControl的列                                          */
            /*****************************反射生成各列*******************************************/
            if (ctl == null || target == null)
                return;
            ctl.PrimaryGrid.Columns.Clear();
            PropertyInfo[] classPropertyInfo;
            classPropertyInfo = target.GetProperties(BindingFlags.Public | BindingFlags.Instance);//获得属性名
            if (classPropertyInfo == null || classPropertyInfo.Length < 1)
                return;
            string catKey = "";
            for (int i = 0; i < classPropertyInfo.Length; i++)
            {
                object[] objs = classPropertyInfo[i].GetCustomAttributes(typeof(DescriptionAttribute), true);
                object[] objs1 = classPropertyInfo[i].GetCustomAttributes(typeof(BrowsableAttribute), true); //根据BrowsableAttribute判断是否显示为列
                bool visible = true;
                if (objs != null && objs.Length > 0)
                {
                    if (objs1 != null && objs1.Length > 0) visible = ((BrowsableAttribute)objs1[0]).Browsable;
                    if (visible)
                    {
                        catKey = ((DescriptionAttribute)objs[0]).Description;
                        if (string.IsNullOrEmpty(catKey))
                            catKey = string.Format("列{0}", i + 1); //设置默认列名
                        GridColumn col = new GridColumn(catKey);
                        col.DataPropertyName = classPropertyInfo[i].Name; //设置绑定的信号属性
                        col.ColumnSortMode = ColumnSortMode.None; //禁止排序
                        col.ResizeMode = ColumnResizeMode.None; //禁止用户调整宽度
                        col.CellStyles.Default.Alignment = DevComponents.DotNetBar.SuperGrid.Style.Alignment.MiddleCenter; //单元格内容居中显示
                        ctl.PrimaryGrid.Columns.Add(col);
                    }
                }
            }
        }

        public static void FullScreenGridControlInWidth(this SuperGridControl ctl)
        {
            /**************************计算公式**********************************************/
            /*                            单元格宽度 = （总可用宽度）/ 列数                               */
            /******************************计算公式******************************************/
            if (ctl == null || ctl.PrimaryGrid.Columns.Count <= 0)
                return;
            int count = ctl.PrimaryGrid.Columns.Count;
            int totalWidth = 0; //可使用的宽度
            if (ctl.VScrollBarVisible)
                totalWidth = ctl.Width - System.Windows.Forms.SystemInformation.VerticalScrollBarWidth;
            else
                totalWidth = ctl.Width;
            if (ctl.PrimaryGrid.ShowRowHeaders)
                totalWidth = totalWidth - ctl.PrimaryGrid.RowHeaderWidth;
            int widthUnit = totalWidth / count;
            for (int i = 0; i < count - 1; i++)
            {
                ctl.PrimaryGrid.Columns[i].Width = widthUnit;
            }
            ctl.PrimaryGrid.Columns[count - 1].Width = totalWidth - (count - 1) * widthUnit; //最后一列占据剩余的宽度
        }

        public static void SaveDataToFile(this SuperGridControl ctl, string filePath)
        {
            if (string.IsNullOrEmpty(filePath))
                return;
            if (ctl == null || ctl.PrimaryGrid.Rows.Count < 1)
                return;
            string dir = Path.GetDirectoryName(filePath);
            if (string.IsNullOrEmpty(dir))
                return;
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);
            ctl.PrimaryGrid.SelectAll();
            ctl.PrimaryGrid.CopySelectedCellsToClipboard(true);
            string str = Clipboard.GetText(TextDataFormat.Text);
            File.AppendAllText(filePath, str);
        }

        public static void SetRowBackColor(this SuperGridControl ctrl, int row)
        {
            if (ctrl == null || ctrl.PrimaryGrid.Rows.Count < 1 || row > ctrl.PrimaryGrid.Rows.Count)
                return;
            int count = ctrl.PrimaryGrid.Rows.Count;
            if (row < 0)
            {
                for (int i = 0; i < count; i++)
                {
                    GridRow currentRow = ctrl.PrimaryGrid.Rows[i] as GridRow;
                    currentRow.CellStyles.Default.Background.Color1 = Color.White;
                    currentRow.CellStyles.Default.Background.Color2 = Color.White;
                }
            }
            else if (row < ctrl.PrimaryGrid.Rows.Count)
            {
                for (int i = 0; i <= row; i++)
                {
                    GridRow currentRow = ctrl.PrimaryGrid.Rows[i] as GridRow;
                    if (i < row)
                    {
                        currentRow.CellStyles.Default.Background.Color1 = Color.LightGray;
                        currentRow.CellStyles.Default.Background.Color2 = Color.LightGray;
                    }
                    else
                    {
                        currentRow.CellStyles.Default.Background.Color1 = Color.Green;
                        currentRow.CellStyles.Default.Background.Color2 = Color.Green;
                    }
                }
            }
            else
            {
                for (int i = 0; i <= row - 1; i++)
                {
                    GridRow currentRow = ctrl.PrimaryGrid.Rows[i] as GridRow;
                    currentRow.CellStyles.Default.Background.Color1 = Color.LightGray;
                    currentRow.CellStyles.Default.Background.Color2 = Color.LightGray;
                }
            }
        }

        /// <summary>
        /// 设置单元格字体颜色
        /// <para>可用于设置执行结果是否成功的指示</para>
        /// </summary>
        /// <param name="ctl">SuperGridControl</param>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <param name="color">颜色</param>
        public static void SetCellForeColor(this SuperGridControl ctl, int row, int col, System.Drawing.Color color)
        {
            if (ctl == null || ctl.PrimaryGrid.Columns.Count < 1 || ctl.PrimaryGrid.Rows.Count < 1)
                return;
            if (row >= ctl.PrimaryGrid.Rows.Count || col >= ctl.PrimaryGrid.Columns.Count)
                return;
            GridRow targetRow = ctl.PrimaryGrid.Rows[row] as GridRow;
            targetRow.Cells[col].CellStyles.Default.TextColor = color;
        }

        /// <summary>
        /// 设置单元格背景色
        /// <para>可用于设置执行结果是否成功的指示</para>
        /// </summary>
        /// <param name="ctl">SuperGridControl</param>
        /// <param name="row">行</param>
        /// <param name="col">列</param>
        /// <param name="color">颜色</param>
        public static void SetCellBackColor(this SuperGridControl ctl, int row, int col, System.Drawing.Color color)
        {
            if (ctl == null || ctl.PrimaryGrid.Columns.Count < 1 || ctl.PrimaryGrid.Rows.Count < 1)
                return;
            if (row >= ctl.PrimaryGrid.Rows.Count || col >= ctl.PrimaryGrid.Columns.Count)
                return;
            GridRow targetRow = ctl.PrimaryGrid.Rows[row] as GridRow;
            targetRow.Cells[col].CellStyles.Default.Background.Color1 = color;
            targetRow.Cells[col].CellStyles.Default.Background.Color2 = color;
        }

        public static void SetRowHeight(this SuperGridControl ctl, int rowHeight, int font)
        {
            if (ctl == null || ctl.PrimaryGrid.Columns.Count <= 0)
                return;
            for (int i = 0; i < ctl.PrimaryGrid.Columns.Count; i++)
            {
                ctl.PrimaryGrid.Columns[i].CellStyles.Default.Font = new System.Drawing.Font("Times New Roman", font);
            }
            ctl.PrimaryGrid.MinRowHeight = rowHeight;
        }
    }
}
