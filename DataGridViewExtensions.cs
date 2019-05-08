using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using System.Reflection;

namespace MyDataGridViewExtensions
{
    public static class DataGridViewExtensions
    {
        /// <summary>
        /// Set BackColor of a cell
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="color"></param>
        public static void SetCellBackColor(this DataGridView dgv, int row, int col, Color color)
        {
            if (row >= dgv.RowCount || col >= dgv.ColumnCount)
                return;
            dgv.Rows[row].Cells[col].Style.BackColor = color;
            dgv.Rows[row].Cells[col].Style.ForeColor = Color.White;
            dgv.Rows[row].Cells[col].ReadOnly = true;
        }

        /// <summary>
        /// Set ForeColor of a cell
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <param name="color"></param>
        public static void SetCellForeColor(this DataGridView dgv, int row, int col, Color color)
        {
            if (row >= dgv.RowCount || col >= dgv.ColumnCount)
                return;
            dgv.Rows[row].Cells[col].Style.ForeColor = color;
            dgv.Rows[row].Cells[col].ReadOnly = true;
        }

        /// <summary>
        /// Reducing flicker, blinking in DataGridView
        /// </summary>
        /// <param name="dgv"></param>
        /// <param name="setting"></param>
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }

        /// <summary>
        /// Set value to multiple cells in a row
        /// </summary>
        /// <param name="dgv">DataGridView</param>
        /// <param name="row">row index</param>
        /// <param name="index">start: column index</param>
        /// <param name="values">Array of object</param>
        public static void SetRangeValuesInRow(this DataGridView dgv, int row, int index, params object[] values)
        {
            if (row >= dgv.RowCount || index >= dgv.ColumnCount || values == null || values.Length < 1)
                return;
            if (index > 0)
            {
                List<object> obj = values.ToList();
                for (int i = 0; i < index; i++)
                {
                    obj.Insert(i, dgv.Rows[row].Cells[i].Value);
                }
                values = obj.ToArray();
            }
            dgv.Rows[row].SetValues(values);
        }
    }
}
