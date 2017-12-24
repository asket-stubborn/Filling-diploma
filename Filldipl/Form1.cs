using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;


namespace Filldipl
{
    public partial class Form1 : Form
    {
        string dpath;
        string opath;
        string dippath;

        
     
        public Form1()
        {
            
            InitializeComponent();
            dpath = "";
            opath = "";
            dippath = "";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog() == DialogResult.OK)
            {
               dpath = dialog.SelectedPath;
                label1.ForeColor = System.Drawing.Color.Green;
                label1.Text = "OK";
                label1.Visible = true;

            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog f1dialog = new OpenFileDialog();
            if (f1dialog.ShowDialog() == DialogResult.OK)
            {
                opath = f1dialog.FileName;
                label2.ForeColor = System.Drawing.Color.Green;
                label2.Text = "OK";
                label2.Visible = true;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog f2dialog = new OpenFileDialog();
            if (f2dialog.ShowDialog() == DialogResult.OK)
            {
                dippath = f2dialog.FileName;
                label3.ForeColor = System.Drawing.Color.Green;
                label3.Text = "OK";
                label3.Visible = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {

            if (label3.Text == label2.Text && label2.Text == label1.Text && label1.Text == "OK")
            {
                Execution exec = new Execution();
                exec.setdippath(dippath);
                exec.setopath(opath);
                exec.bustfolder(dpath);
                MessageBox.Show("Выполнено!", "Выполнено!", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }
            else MessageBox.Show("Ошибка!", "Не все параметры введены!", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

       
    }
    public class Execution
    {
        private string opathe;
        private string dippathe;



        string catnametrue;
        string breedtrue;
        string colortrue;
        string classofcattrue;
        string sexofcattrue;
        string birthtrue;
        string ownertrue;
        string pedigreetrue;

        public Execution()
        {
            catnametrue="";
            breedtrue = "";
            colortrue = "";
            classofcattrue = "";
            sexofcattrue = "";
            birthtrue = "";
            ownertrue = "";
            pedigreetrue = "";
        }
        public void setopath(string opathf)        {
            opathe = opathf;
        }

        public void setdippath(string dirpathf)        {
            dippathe = dirpathf;
        }
        // вот это всё надо доделать. Внутренние функции кнопки заполнения

        
        public int bustfolder(string dpathin) {
            int err=0;
            int readerr = 0;
            int writedatatooerr = 0;
            int writedatatoderr = 0;
            string diro = dpathin + "\\" + "Оценочные_листы";
            string dird = dpathin + "\\" + "Дипломы";
            Regex rgx = new Regex(@"((doc)||(docx))$");
            if (!Directory.Exists(diro)) { Directory.CreateDirectory(diro); }
            if (!Directory.Exists(dird)) { Directory.CreateDirectory(dird); }
            string[] files = Directory.GetFiles(dpathin);

            foreach (string file in files)
            { if (rgx.IsMatch(file)) {
                    readerr = readingdata(file,dpathin); // обработка ошибок на вход
                    if (readerr == 0) {
                        if (catnametrue == "" || breedtrue == "" || colortrue == "" || classofcattrue == "" || sexofcattrue == "" || birthtrue == "" || ownertrue == "" || pedigreetrue == "")
                            MessageBox.Show("Предупреждение",file+" не содержит всех нужных данных", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
                        writedatatooerr = pushdataintoo(diro,opathe);
                        writedatatoderr = pushdataintodip(dird,dippathe);
                          }
                    catnametrue="";
                    breedtrue = "";
                    colortrue = "";
                    classofcattrue = "";
                    sexofcattrue = "";
                    birthtrue = "";
                    ownertrue = "";
                    pedigreetrue = "";

                }
            } 
            return err;
        }
        public int readingdata(string file,string dpathin)
        {
            int err = 0;
            Regex catname = new Regex(@"((К((личка)|(ЛИЧКА)))|(N((ame)|(AME))))");
            Regex breed = new Regex(@"((П((орода)|(ОРОДА)))|(B((reed)|(REED)))|(R((asse)|(ASSE))))");
            Regex color = new Regex(@"((О((крас)|(КРАС)))|(C((olor)|(OLOR))))");
            Regex classofcat = new Regex(@"((К((ласс)|(ЛАСС)))|(C((LASS)|(lass)))|(K((asse)|(ASSE))))");
            Regex sexofcat = new Regex(@"((П((ол)|(ОЛ)))|(S((EX)|(ex)))|(G((eshlecht)|(ESHLECHT))))");
            Regex birth = new Regex(@"((Д((ата)|(АТА)))|(B((irth)|(IRTH)))|(G((eboren)|(EBOREN)))))");
            Regex owner = new Regex(@"((В((ладелец)|(ЛАДЕЛЕЦ)))|(O((wner)|(WNER)))|E((igentümer)|(IGENTÜMER)))))");
            Regex pedigree = new Regex(@"((((Р)|(р))|((одословн)|(одословн)))|(((P)|(p)))|((edigree)|(EDIGREE)))");

            Word.Application word = new Word.Application();
            object missing = Type.Missing;
            object filename = dpathin+"\\"+ file;
            Word.Document doc = word.Documents.Open(ref filename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            for (int j=1; j<doc.Tables.Count;j++)

            for (int i = 1; i < doc.Tables[2].Rows.Count; i++)
            {
                
                
                    if (catname.IsMatch(doc.Tables[j].Cell(i, 1).Range.Text)) { catnametrue = doc.Tables[j].Cell(i, 2).Range.Text; catnametrue = catnametrue.Replace("\r\a", string.Empty); }
                    if (breed.IsMatch(doc.Tables[j].Cell(i, 1).Range.Text)) { breedtrue = doc.Tables[j].Cell(i, 2).Range.Text; breedtrue = breedtrue.Replace("\r\a", string.Empty); }
                    if (color.IsMatch(doc.Tables[j].Cell(i, 1).Range.Text)) { colortrue = doc.Tables[j].Cell(i, 2).Range.Text; colortrue = breedtrue.Replace("\r\a", string.Empty); }
                    if (sexofcat.IsMatch(doc.Tables[j].Cell(i, 1).Range.Text)) { sexofcattrue = doc.Tables[j].Cell(i, 2).Range.Text; sexofcattrue = sexofcattrue.Replace("\r\a", string.Empty); }
                    if (birth.IsMatch(doc.Tables[j].Cell(i, 1).Range.Text)) { birthtrue = doc.Tables[j].Cell(i, 2).Range.Text; birthtrue = birthtrue.Replace("\r\a", string.Empty); }
                    if (owner.IsMatch(doc.Tables[j].Cell(i, 1).Range.Text)) { ownertrue = doc.Tables[j].Cell(i, 2).Range.Text; ownertrue = ownertrue.Replace("\r\a", string.Empty); }
                    if (pedigree.IsMatch(doc.Tables[j].Cell(i, 1).Range.Text)) { pedigreetrue = doc.Tables[j].Cell(i, 2).Range.Text; pedigreetrue = pedigreetrue.Replace("\r\a", string.Empty); }
                }
            doc.Close();
            return err;
        }
        public int pushdataintoo(string diro, string opathe) {
            int err = 0;

            Word.Application word = new Word.Application();
            object missing = Type.Missing;
            object filenamet = opathe;
            Word.Document dot = word.Documents.Open(ref filenamet, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            string filename = string.Join("_", ownertrue.Split(Path.GetInvalidFileNameChars()));
            object fullfilename = diro +"//" + filename;
            
            dot.SaveAs2(ref fullfilename);

            dot.Close();

            Word.Document doco = word.Documents.Open(ref fullfilename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);


            doco.Bookmarks["Birth"].Range.Text = birthtrue;
            doco.Bookmarks["Breed"].Range.Text = breedtrue;
            doco.Bookmarks["Class"].Range.Text = classofcattrue;
            doco.Bookmarks["Color"].Range.Text = colortrue;
            doco.Bookmarks["Sex"].Range.Text = sexofcattrue;

            doco.Close();

            return err;

        }
        public int pushdataintodip(string dird, string dippathe) {
            int err = 0;

            Word.Application word = new Word.Application();
            object missing = Type.Missing;
            object filenamet = dippathe;
            Word.Document dot = word.Documents.Open(ref filenamet, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            string filename = string.Join("_", ownertrue.Split(Path.GetInvalidFileNameChars()));
            object fullfilename = dird + "//" + filename;

            dot.SaveAs2(ref fullfilename);

            dot.Close();

            Word.Document docdip = word.Documents.Open(ref fullfilename, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);


            docdip.Bookmarks["Birth"].Range.Text = birthtrue;
            docdip.Bookmarks["Breed"].Range.Text = breedtrue;
            docdip.Bookmarks["Catname"].Range.Text = catnametrue;
            docdip.Bookmarks["Class"].Range.Text = classofcattrue;
            docdip.Bookmarks["Color"].Range.Text = colortrue;
            docdip.Bookmarks["Owner"].Range.Text = ownertrue;
            docdip.Bookmarks["Pedigree"].Range.Text = pedigreetrue;
            docdip.Bookmarks["Sex"].Range.Text = sexofcattrue;

            docdip.Close();


            return err;
        }

       
    }
}
