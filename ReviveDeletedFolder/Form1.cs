using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using SHDocVw;
using Shell32;

namespace ReviveDeletedFolder
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            timer1.Interval = 60000; // 1分
            //timer1.Interval = 1000;
            timer1.Start();
        }

        private void Form1_Load( object sender, EventArgs e )
        {
            new Thread( new ThreadStart( HideForm ) ).Start();
        }

        private delegate void HideFormDelegate();

        private void HideForm()
        {
            if ( InvokeRequired )
            {
                Invoke( new HideFormDelegate( HideForm ) );
                return;
            }
            this.Hide();
        }

        private void 終了ToolStripMenuItem_Click( object sender, EventArgs e )
        {
            ApplicationExit();
        }


        private void ApplicationExit()
        {
            m_AppExit = true;

            timer1.Stop();
            notifyIcon1.Visible = false;
            Application.Exit();
        }
        private bool m_AppExit = false;

        private void timer1_Tick( object sender, EventArgs e )
        {
            var livingDirectories = new List<DirectoryManager>();

            //COMのShellクラス作成
            Shell shell = new Shell();

            //IEとエクスプローラの一覧を取得
            ShellWindows win = shell.Windows();

            foreach ( IWebBrowser2 web in win )
            {
                try
                {
                    //エクスプローラのみ(IEを除外)
                    if ( Path.GetFileName( web.FullName ).ToUpper() == "EXPLORER.EXE" )
                    {
                        string localPath = new Uri( web.LocationURL ).LocalPath;

                        if ( !livingDirectories.Any( q => q.Name == web.LocationName ) )
                            livingDirectories.Add( new DirectoryManager( web.LocationName, localPath ) );
                    }
                }
                catch { }
            }

            // 死んでいったフォルダー達を弔う
            foreach ( var lived in m_LivedDirectories )
            {
                if ( !livingDirectories.Any( q => q.Name == lived.Name ) )
                {
                    var dir = m_DeletedDirectories.FirstOrDefault(q => q.Name == lived.Name);
                    if ( dir != null )
                        m_DeletedDirectories.Remove( dir );

                    m_DeletedDirectories.Insert( 0, new DirectoryManager( lived ) );
                }
            }

            m_LivedDirectories = livingDirectories;
        }

        private List<DirectoryManager> m_LivedDirectories = new List<DirectoryManager>();
        private List<DirectoryManager> m_DeletedDirectories = new List<DirectoryManager>();

        private void 表示ToolStripMenuItem_Click( object sender, EventArgs e )
        {
            表示();
        }

        private void notifyIcon1_DoubleClick( object sender, EventArgs e )
        {
            表示();
        }

        private void 表示()
        {
            listView1.Items.Clear();

            foreach ( var dir in m_DeletedDirectories )
                listView1.Items.Add( new ListViewItem( new string[] { dir.Name, dir.Path } ) );

            int ScreenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int Screenheigth = Screen.PrimaryScreen.WorkingArea.Height;
            int AppWidth = this.Width;
            int AppHeight = this.Height;
            int AppLeftXPos = ScreenWidth - AppWidth;
            int AppLeftYPos = Screenheigth - AppHeight;
            Rectangle tempRect = new Rectangle( AppLeftXPos, AppLeftYPos, AppWidth, AppHeight );
            this.DesktopBounds = tempRect;

            this.Show();
        }

        internal class DirectoryManager
        {
            public DirectoryManager(string name, string path)
            {
                Name = name;
                Path = path;
            }

            public DirectoryManager( DirectoryManager dir )
                : this( dir.Name, dir.Path )
            {
            }

            public string Name { get; set; }
            public string Path { get; set; }
        }

        private void Form1_FormClosing( object sender, FormClosingEventArgs e )
        {
            if ( !m_AppExit )
            {
                e.Cancel = true;
                this.Hide();
            }
        }

        private void listView1_DoubleClick( object sender, EventArgs e )        
        {
            ListView lvw = (ListView)sender;
            if ( lvw != null )
            {
                var item = lvw.SelectedItems[0];
                var ss = item.SubItems[1] as System.Windows.Forms.ListViewItem.ListViewSubItem;
                string path = ss.Text;
                if ( Directory.Exists( path ) )
                {
                    try { System.Diagnostics.Process.Start( path ); }
                    catch
                    {
                        MessageBox.Show("フォルダーを開けませんでした。", "情報", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void listView1_KeyDown( object sender, KeyEventArgs e )
        {
            switch(e.KeyData)
            {
                case Keys.F5:
                    表示();
                    break;
            }
        }
    }
}
