using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace HSS
{
    /// <summary>
    /// Enterキーが押された時に、Tabキーが押されたのと同じ動作をする
    /// （現在のセルを隣のセルに移動する）DataGridView
    /// </summary>
    public class DataGridViewEx : DataGridView
    {
        [System.Security.Permissions.UIPermission(
            System.Security.Permissions.SecurityAction.LinkDemand,
            Window = System.Security.Permissions.UIPermissionWindow.AllWindows)]
        protected override bool ProcessDialogKey(Keys keyData)
        {
            // ①Enterキー、→キーが押された時は、Tabキーが押されたようにする
            // ②ReadOnlyのセルはスキップする
            if ((keyData & Keys.KeyCode) == Keys.Enter || (keyData & Keys.KeyCode) == Keys.Right)
            {
                // 次のセルがReadonlyのときReadOnlyではないセルまで検証
                if (this.Columns[this.CurrentCell.ColumnIndex + 1].ReadOnly)
                {
                    int Col = this.CurrentCell.ColumnIndex + 1;
                    while (Col < this.Columns.Count)
                    {
                        if (this.Columns[Col].ReadOnly == false)
                        {
                            break;
                        }

                        Col++;
                    }

                    // 同じ行内での移動
                    if (Col < this.Columns.Count)
                    {
                        this.CurrentCell = this[this.Columns[Col - 1].Name, this.CurrentRow.Index];
                    }
                    // 次の行へ移動（ダイレクトにカラム[1]へ遷移）
                    else if (this.CurrentRow.Index < this.Rows.Count - 1)
                    {
                        this.CurrentCell = this[this.Columns[1].Name, this.CurrentRow.Index + 1];
                    }

                    //this.CurrentCell = this[this.Columns[this.CurrentCell.ColumnIndex + 1].Name, this.CurrentRow.Index];
                }
                return this.ProcessTabKey(keyData);
            }
            return base.ProcessDialogKey(keyData);
        }

        [System.Security.Permissions.SecurityPermission(
            System.Security.Permissions.SecurityAction.LinkDemand,
            Flags = System.Security.Permissions.SecurityPermissionFlag.UnmanagedCode)]
        protected override bool ProcessDataGridViewKey(KeyEventArgs e)
        {
            // ①Enterキー、→キーが押された時は、Tabキーが押されたようにする
            // ②ReadOnlyのセルはスキップする
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Right)
            {
                // 次のセルがReadonlyのときReadOnlyではないセルまで検証
                if (this.Columns[this.CurrentCell.ColumnIndex + 1].ReadOnly)
                {
                    int Col = this.CurrentCell.ColumnIndex + 1;
                    while (Col < this.Columns.Count)
                    {
                        if (this.Columns[Col].ReadOnly == false)
                        {
                            break;
                        }

                        Col++;
                    }

                    // 同じ行内での移動
                    if (Col < this.Columns.Count)
                    {
                        this.CurrentCell = this[this.Columns[Col - 1].Name, this.CurrentRow.Index];
                    }
                    else if (this.CurrentRow.Index < this.Rows.Count - 1)
                    {
                        // 次の行へ移動（ダイレクトにカラム[1]へ遷移）
                        this.CurrentCell = this[this.Columns[1].Name, this.CurrentRow.Index + 1];
                    }
                }
                return this.ProcessTabKey(e.KeyCode);
            }
            return base.ProcessDataGridViewKey(e);
        }
    }
}
