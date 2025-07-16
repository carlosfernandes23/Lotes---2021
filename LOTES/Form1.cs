using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using TSM = Tekla.Structures.Model;
using Tekla.Structures;
using Tekla.Structures.Geometry3d;
using Tekla.Structures.Model.UI;
using Tekla.Structures.Filtering;
using Tekla.Structures.Filtering.Categories;
using System.Collections;
using Tekla.Structures.Model;


namespace LOTES
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string path = null;
        string OBRA = null;
        List<string> lote;
        List<int> lotes=new List<int>();
        String[] allfiles;

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.SelectedIndex = 1;
            comboBox3.SelectedIndex = 1;
            comboBox7.SelectedIndex = 0;
        }

        void refreshprog()
        {

            TopMost = true;

            TSM.Model M = new Model();

            try
            {
                string teste = M.GetInfo().ModelName.Replace(".db1", "");
                path = M.GetInfo().ModelPath.Replace(teste, "");
                OBRA = M.GetProjectInfo().ProjectNumber;
            }
            catch (Exception ex)
            {
                MessageBox.Show("ERRO P.F. ABRA O TEKLA SE ESTIVER ABERTO P.F. REINICIE ESTE PROGRAMA" + ex.Message, "Erro", MessageBoxButtons.OK, MessageBoxIcon.Error);
                this.Close();

            }
        }

        void ACTULIZALOTES()
        {
            // procurar ficheiros de texto 
            allfiles = Directory.GetFiles(path, "OFelizLotes.txt", System.IO.SearchOption.AllDirectories);
            // ler ficheiros
            lote = new List<string>();

            foreach (var file in allfiles)
            {
                StreamReader leitor = new StreamReader(file);
                while (!leitor.EndOfStream)
                {
                    lote.Add(leitor.ReadLine());
                }
                leitor.Close();
            }
            lote.Sort();
               foreach (string item in lote.Distinct().ToList())
            {
                try
                {
                    lotes.Add(int.Parse(item));
                }
                catch (Exception)
                {

                   
                }
             
            }
            lotes.Sort();
            List<string> lotestr = new List<string>();
            foreach (int item in lotes.Distinct().ToList())
            {
                lotestr.Add(item.ToString());
            }


            foreach (var file in allfiles)
            {
                File.Delete(file);
             
                File.WriteAllLines(file, lotestr);
            }
            comboBox2.Items.Clear();
            for (int i = 1; i <= 250; i++)
            {
                bool contem = false;
                if (lote != null)
                {
                    foreach (var item in lote)
                    {
                        if (i.ToString() == item.Replace(" ", ""))
                        {
                            contem = true;
                        }
                    }
                }
                if (contem == false)
                {
                    comboBox2.Items.Add(i);
                }
            }
            comboBox2.SelectedIndex = 0;
            if (allfiles.Length == 0)
            {
                if (!Directory.Exists(path + "\\logs"))
                {
                    Directory.CreateDirectory(path + "\\logs");
                }
            }
            listBox1.DataSource = lotes;
        }

        void ACTULIZAlist()
        {
            // procurar ficheiros de texto 
            allfiles = Directory.GetFiles(path, "OFelizLotes.txt", System.IO.SearchOption.AllDirectories);
            // ler ficheiros
            lote = new List<string>();

            foreach (var file in allfiles)
            {
                StreamReader leitor = new StreamReader(file);
                while (!leitor.EndOfStream)
                {
                    lote.Add(leitor.ReadLine());
                }
                leitor.Close();
            }
            
           
            foreach (string item in lote.Distinct().ToList())
            {
                lotes.Add(int.Parse(item));
            }
            lotes.Sort();
            List<string> lotestr = new List<string>();
            foreach (int item in lotes.Distinct().ToList())
            {
                lotestr.Add(item.ToString());
            }


            foreach (var file in allfiles)
            {
                File.Delete(file);
                File.WriteAllLines(file, lotestr);
            }
            if (allfiles.Length == 0)
            {
                if (!Directory.Exists(path + "\\logs"))
                {
                    Directory.CreateDirectory(path + "\\logs");
                }
            }
            listBox1.DataSource = lotestr;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            TSM.Model M = new TSM.Model();
            if (M.GetProjectInfo().ProjectNumber == label10.Text)
            {

                //CRIAR FILTRO 1
                string MATERIALFILTRAR = null;
                if (textBox1.Text.Replace(" ", "") == "")
                {
                    MATERIALFILTRAR = "*";
                }
                else
                {
                    MATERIALFILTRAR = textBox1.Text;
                }
                PartFilterExpressions.Material MATERIAL = new PartFilterExpressions.Material();
                StringConstantFilterExpression FILTROMATERIAL = new StringConstantFilterExpression(MATERIALFILTRAR);
                BinaryFilterExpression FILTRO1 = new BinaryFilterExpression(MATERIAL, StringOperatorType.IS_EQUAL, FILTROMATERIAL);

                //CRIAR FILTRO 2
                string PERFILFILTRAR = null;
                if (textBox2.Text.Replace(" ", "") == "")
                {
                    PERFILFILTRAR = "*";
                }
                else
                {
                    PERFILFILTRAR = textBox2.Text;
                }
                PartFilterExpressions.Profile PERFIL = new PartFilterExpressions.Profile();
                StringConstantFilterExpression FILTROPERFIL = new StringConstantFilterExpression(PERFILFILTRAR);
                BinaryFilterExpression FILTRO2 = new BinaryFilterExpression(PERFIL, StringOperatorType.IS_EQUAL, FILTROPERFIL);

                //CRIAR FILTRO 3
                string FASEFILTRAR = null;
                if (textBox3.Text.Replace(" ", "") == "")
                {
                    FASEFILTRAR = "*";
                }
                else
                {
                    FASEFILTRAR = textBox3.Text;
                }

                BinaryFilterExpression FILTRO3 = new BinaryFilterExpression(new AssemblyFilterExpressions.Phase(), StringOperatorType.IS_EQUAL, new StringConstantFilterExpression(FASEFILTRAR));

                //CRIAR FILTRO 4
                string LOTEFILTRAR = null;
                if (comboBox1.Text == "NÃO")
                {
                    LOTEFILTRAR = "*";
                }
                else
                {
                    LOTEFILTRAR = "0";
                }

                BinaryFilterExpression FILTRO4 = new BinaryFilterExpression(new TemplateFilterExpressions.CustomString("USERDEFINED.lote_number"), StringOperatorType.IS_EQUAL, new StringConstantFilterExpression(LOTEFILTRAR));


                // CRIAR GRUPO DE FILTROS
                BinaryFilterExpressionCollection A = new BinaryFilterExpressionCollection();


                A.Add(new BinaryFilterExpressionItem(FILTRO1));
                A.Add(new BinaryFilterExpressionItem(FILTRO2));
                A.Add(new BinaryFilterExpressionItem(FILTRO3));
                A.Add(new BinaryFilterExpressionItem(FILTRO4));

                //CRIAR FICHEIRO DO GRUPO DE FILTROS 
                string AttributesPath = M.GetInfo().ModelPath + @"\attributes";
                string FilterName = Path.Combine(AttributesPath, "filter");
                Filter Filter = new Filter(A);
                Filter.CreateFile(FilterExpressionFileType.OBJECT_GROUP_VIEW, FilterName);

                ModelViewEnumerator ViewEnum1 = ViewHandler.GetAllViews();
                while (ViewEnum1.MoveNext())
                {
                    if (ViewEnum1.Current.Name == "DIRECCAOOBRA")
                    {
                        ViewEnum1.Current.Delete();
                    }
                }

                //CRIAR VISTA
                TSM.UI.View View = new TSM.UI.View();
                ViewVisibilitySettings n = new ViewVisibilitySettings();
                n.BoltHolesVisible = true;
                n.BoltsVisible = true;
                n.WeldsVisible = false;
                n.ConstructionLinesVisible = false;
                n.CutsVisible = false;
                n.CutsVisibleInComponents = false;
                n.ReferenceObjectsVisible = false;
                n.ComponentsVisibleInComponents = false;
                n.FittingsVisible = false;
                n.FittingsVisibleInComponents = false;
                n.ReferenceObjectsVisible = false;
                n.PourBreaksVisible = false;
                n.SurfaceTreatmentsVisible = false;
                n.PartsVisibleInComponents = true;
                n.LoadsVisible = false;
                n.SurfaceTreatmentsVisible = false;
                n.ComponentsVisible = false;
                n.ComponentsVisibleInComponents = false;

                View.VisibilitySettings =n;        
                View.Name = "DIRECCAOOBRA";
                View.ViewCoordinateSystem.AxisX = new Vector(1, 0, 0);
                View.ViewCoordinateSystem.AxisY = new Vector(0, 1, 0);
                // Work area has to be set for new views
                View.WorkArea.MinPoint = new Tekla.Structures.Geometry3d.Point(-3000, -3000, -3000);
                View.WorkArea.MaxPoint = new Tekla.Structures.Geometry3d.Point(15000, 33000, 30000);
                View.ViewDepthUp = 50000;
                View.ViewDepthDown = 2000;
                View.ViewFilter = "filter";
                
                View.Insert();
                TeklaStructures.Connect();
                ModelViewEnumerator ViewEnum = ViewHandler.GetAllViews();
                while (ViewEnum.MoveNext())
                {
                    ViewHandler.HideView(ViewEnum.Current);
                }
                ViewHandler.ShowView(View);
                View.Modify();
                TeklaStructures.ExecuteScript("akit.Callback(\"acmd_fit_workarea\", \"\", \"main_frame\");");
            }

        }

        internal bool verificaçaopintura(ArrayList PECAS)
        {
            bool aprovado = true;
            foreach (TSM.Part item in PECAS)
            {
                string r = null;
                item.GetAssembly().GetReportProperty("pintura", ref r);

                if (r==""||r==null)
                {
                    aprovado = false;
                }
            }

            return aprovado;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TSM.Model Model = new TSM.Model();
            if (Model.GetProjectInfo().ProjectNumber == label10.Text)
            {
                ArrayList PECASRAL = SELECTMODEL();
                string[] profilePrefixes = { "AT", "TC", "FW", "IW", "IC", "MT", "TL", "P0", "P1", "P2", "P4", "P5", "P6" };

                foreach (TSM.Part item in PECASRAL)
                {
                    string profile = null;
                    item.GetReportProperty("PROFILE", ref profile);

                    string material = null;
                    item.GetReportProperty("MATERIAL", ref material);

                    bool perfilMatch = !string.IsNullOrEmpty(profile) &&
                        profilePrefixes.Any(p => profile.StartsWith(p, StringComparison.OrdinalIgnoreCase));
                    bool materialMatch = !string.IsNullOrEmpty(material) &&
                        (material.IndexOf("DX51", StringComparison.OrdinalIgnoreCase) >= 0 ||
                         material.IndexOf("S220GD", StringComparison.OrdinalIgnoreCase) >= 0);

                    if (perfilMatch || materialMatch)
                    {
                        string chapa = null;
                        item.GetUserProperty("CHAPA_LACADA", ref chapa);
                        if (string.IsNullOrEmpty(chapa))
                        {
                            MessageBox.Show(
                                "O perfil \"" + profile + "\" com material \"" + material + "\" não tem RAL atribuído." + Environment.NewLine +
                                "Por favor atribua o RAL a este perfil.",
                                "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }

            if (Model.GetProjectInfo().ProjectNumber == label10.Text)
            {
                if (!textBox4.Text.Contains(";"))
                {

               
                ArrayList PECAS = SELECTMODEL();

                if (verificaçaopintura(PECAS))
                {
                    string existe = null;
                    bool x = false;
                    bool c = false;

                    foreach (TSM.Part item in PECAS)
                    {
                        item.GetUserProperty("lote_number", ref existe);


                        if (!String.IsNullOrEmpty(existe) && x == false)
                        {
                            x = true;
                            DialogResult d = MessageBox.Show("Foram detectadas peças ja loteadas deseja sobrepôr" + Environment.NewLine + Environment.NewLine + "Atenção este recurso pode surgir erros recomenda-se anular o lote e atribuir-lo novamente." + Environment.NewLine + Environment.NewLine + "Pode utilizar a mesma seleção apenas terá de carregar no botão apaga lote e carregar novamente em adicionar lote.", "Atenção", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation);
                            if (d == DialogResult.Yes)
                            {
                                c = false;
                            }
                            else
                            {
                                c = true;
                            }
                        }
                        else
                        {
                            if (c == false)
                            {
                                item.SetUserProperty("lote_number", comboBox2.Text);
                                item.SetUserProperty("lote_data", dateTimePicker1.Value.ToShortDateString());
                                item.SetUserProperty("Local_descarga", textBox4.Text);
                            }
                        }
                    }
                    TSM.Model M = new Model();

                    File.AppendAllText(M.GetInfo().ModelPath.ToString() + "\\logs\\OFelizLotes.txt", comboBox2.Text + Environment.NewLine);
                    ACTULIZAlist();

                    lblinfo.Text = "Concluido";
                    Model.CommitChanges();
                }
                else
                {
                    MessageBox.Show("Existem conjuntos sem esquema" + Environment.NewLine + "Não é permitido lotear sem esquema atribuido","ERRO",MessageBoxButtons.OK,MessageBoxIcon.Error);
                }
                }else
                {
                    MessageBox.Show("Utilize outro caracter no local de descarga que nao seja o ';'" + Environment.NewLine + "Não é permitido lotear.", "ERRO", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        internal ArrayList SELECTMODEL()
        {
            lblinfo.Text = "Comunicar com tekla";
            ArrayList segParts = new ArrayList();
            ArrayList assembly = new ArrayList();
            int a = 0;

            TSM.Model model = new TSM.Model();
            TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();

            while (modelEnumerator.MoveNext())
            {
                TSM.Part part = modelEnumerator.Current as TSM.Part;
                TSM.Assembly ass = modelEnumerator.Current as TSM.Assembly;

                if (ass != null)
                {
                    foreach (TSM.Part item in ass.GetSecondaries())
                    {
                        segParts.Add(item);
                        a++;
                        lblinfo.Text = "procurar peças " + a + " encontradas";
                    }
                    segParts.Add(ass.GetMainPart());
                    a++;
                    lblinfo.Text = "procurar peças " + a + " encontradas";
                }
            }


            if (segParts.Count == 0)
            {
                TSM.ModelObjectEnumerator moe = new TSM.UI.ModelObjectSelector().GetSelectedObjects();
                moe.SelectInstances = false;
                while (moe.MoveNext())
                {
                    TSM.Part myPart = moe.Current as TSM.Part;

                    if (myPart != null)
                    {
                        TSM.Assembly ass = myPart.GetAssembly();
                        string asse = myPart.GetAssembly().Identifier.GUID.ToString();
                        if (!assembly.Contains(asse))
                        {
                            assembly.Add(asse);
                            foreach (var item in ass.GetSecondaries())
                            {
                                segParts.Add(item);
                                a++;
                                lblinfo.Text = "Procurar peças " + a + " encontradas";
                            }
                            segParts.Add(ass.GetMainPart());
                            a++;
                            lblinfo.Text = "Procurar peças " + a + " encontradas";
                        }
                    }
                }
            }

            lblinfo.Text = "" + segParts.Count + " Peças encontradas";
            assembly.Clear();
            return segParts;

        }

        protected void button3_Click(object sender, EventArgs e)
        {

            //TSM.Model MyModel = new TSM.Model();


            //BinaryFilterExpression FilterExpression1 = new BinaryFilterExpression(new ObjectFilterExpressions.Type(), NumericOperatorType.IS_EQUAL,
            //    new NumericConstantFilterExpression(TeklaStructuresDatabaseTypeEnum.ASSEMBLY));

            //BinaryFilterExpressionCollection FilterDefinition = new BinaryFilterExpressionCollection { new BinaryFilterExpressionItem(FilterExpression1, BinaryFilterOperatorType.BOOLEAN_AND) };

            //ModelObjectEnumerator modelEnumerator = MyModel.GetModelObjectSelector().GetObjectsByFilter(FilterDefinition);

            //lblinfo.Text = "Comunicar com tekla";
            //ArrayList segParts = new ArrayList();
            //int a = 0;
            //while (modelEnumerator.MoveNext())
            //{
            //    TSM.Part part = modelEnumerator.Current as TSM.Part;
            //    TSM.Assembly ass = modelEnumerator.Current as TSM.Assembly;
            //    string lote = null;

            //    if (ass != null)
            //    {
            //        ass.GetMainPart().GetReportProperty("lote_number", ref lote);
            //        if (!segParts.Contains(lote))
            //        {
            //            segParts.Add(lote);
            //        }
            //        a++;
            //        lblinfo.Text = "procurar peças " + a + " encontradas";
            //    }
            //}
            //lbl_lote.Text = (segParts.Count + 1).ToString();
        }

        private void TXTPESO_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar) &&
                (e.KeyChar != '.'))
            {
                e.Handled = true;
            }

            // only allow one decimal point
            if ((e.KeyChar == '.') && ((sender as TextBox).Text.IndexOf('.') > -1))
            {
                e.Handled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            ArrayList PECAS = SELECTMODEL1();

            double peso = 0;
            double somapeso = 0;
            if (PECAS.Count != 0)
            {


                foreach (Part item in PECAS)
                {
                    item.GetReportProperty("PROFILE_WEIGHT_NET", ref peso);
                    somapeso += peso;
                }
                Pesoselec.Text = somapeso.ToString("0");
                txtresultado.Text = "Restam " + (double.Parse(TXTPESO.Text) - somapeso).ToString("0") + " Kg";

                if (somapeso >= double.Parse(TXTPESO.Text))
                {
                    txtresultado.BackColor = System.Drawing.Color.Red;
                }
                else
                {
                    txtresultado.BackColor = TXTPESO.BackColor;
                }
            }
        }

        internal ArrayList SELECTMODEL1()
        {
            lblinfo.Text = "Comunicar com tekla";
            ArrayList segParts = new ArrayList();
            ArrayList assembly = new ArrayList();
            int a = 0;


            TSM.Model model = new TSM.Model();
            if (model.GetProjectInfo().ProjectNumber == label10.Text)
            {
                TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();

                while (modelEnumerator.MoveNext())
                {
                    TSM.Part part = modelEnumerator.Current as TSM.Part;
                    TSM.Assembly ass = modelEnumerator.Current as TSM.Assembly;

                    if (ass != null)
                    {
                        foreach (TSM.Part item in ass.GetSecondaries())
                        {
                            segParts.Add(item);
                            a++;
                            lblinfo.Text = "procurar peças " + a + " encontradas";
                        }
                        segParts.Add(ass.GetMainPart());
                        a++;
                        lblinfo.Text = "procurar peças " + a + " encontradas";
                    }
                }
                if (segParts.Count == 0)
                {
                    TSM.ModelObjectEnumerator moe = new TSM.UI.ModelObjectSelector().GetSelectedObjects();
                    moe.SelectInstances = false;
                    while (moe.MoveNext())
                    {
                        TSM.Part myPart = moe.Current as TSM.Part;

                        if (myPart != null)
                        {
                            TSM.Assembly ass = myPart.GetAssembly();
                            string asse = myPart.GetAssembly().Identifier.GUID.ToString();
                            if (!assembly.Contains(asse))
                            {
                                assembly.Add(asse);
                                foreach (var item in ass.GetSecondaries())
                                {
                                    segParts.Add(item);
                                    a++;
                                    lblinfo.Text = "Procurar peças " + a + " encontradas";
                                }
                                segParts.Add(ass.GetMainPart());
                                a++;
                                lblinfo.Text = "Procurar peças " + a + " encontradas";
                            }
                        }
                    }
                }


                lblinfo.Text = "" + segParts.Count + " Peças encontradas";
                assembly.Clear();
            }
            return segParts;

        }

        private void comboBox2_SelectedValueChanged(object sender, EventArgs e)
        {
            panel2.Focus();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            comboBox4.SelectedIndex = 0;
            refreshprog();
            ACTULIZALOTES();
            label10.Text = OBRA;
        }

        private void listBox1_DoubleClick(object sender, EventArgs e)
        {
            List<string>lotes = Getallpartswhithfilter("lote_number", listBox1.SelectedItem.ToString());

            if (lotes.Count>1)
            {
                string message = "Atenção existe mais do que um esquema neste lote apenas o primeiro será carregado."+Environment.NewLine+Environment.NewLine;
                foreach (string item in lotes)
                {
                    message += item + Environment.NewLine;
                }


                MessageBox.Show(message, "ATENÇÃO", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            try
            {
                comboBox2.Text = lotes[0].Split(';')[0].Trim();
            }
            catch (Exception) { }
            try
            {
                dateTimePicker1.Text = lotes[0].Split(';')[1].Trim();
            }
            catch (Exception) { }
            try
            {
                comboBox7.Text = lotes[0].Split(';')[2].Trim();
            }
            catch (Exception) { }
            try
            {
                comboBox6.Text = lotes[0].Split(';')[3].Trim();
            }
            catch (Exception) { }
            try
            {
                textBox4.Text = lotes[0].Split(';')[5].Trim();
            }
            catch (Exception) { }

        }

        List<string> Getallpartswhithfilter(string atributo, string parametro)
        {
            string lote_number = null;
            string lote_data = null;
            string Local_descarga = null;
            string pintura = null;
            string OperacaoFabrica = null;
            List<String> PINTURA = new List<string>(); 
            TSM.Model m = new Model();
            if (m.GetProjectInfo().ProjectNumber == label10.Text)
            {
                BinaryFilterExpression filterExpression = FAZFILTROS(atributo, parametro);


                TSM.ModelObjectEnumerator modelObjectSelector = m.GetModelObjectSelector().GetObjectsByFilter(filterExpression);
                ArrayList ojectos = new ArrayList();

                foreach (var item in modelObjectSelector)
                {
                    ojectos.Add(item);

                    if (item.ToString().Contains("Beam")|| item.ToString().Contains("ContourPlate"))
                    {
                     Part Parte = item as Part;

                        Parte.GetUserProperty("lote_number", ref lote_number);
                        Parte.GetUserProperty("lote_data", ref lote_data);
                        Parte.GetAssembly().GetUserProperty("pintura", ref pintura);
                        Parte.GetAssembly().GetUserProperty("OperacaoFabrica", ref OperacaoFabrica);
                        Parte.GetUserProperty("Local_descarga", ref Local_descarga);
                      
                    }
                    PINTURA.Add(lote_number + ";" + lote_data + ";" + pintura + ";" + OperacaoFabrica + ";" + Local_descarga);
                
                }
                TSM.UI.ModelObjectSelector SEL = new TSM.UI.ModelObjectSelector();
                SEL.Select(ojectos);


            }
            List<String> PINTURAS = PINTURA.Distinct().ToList();

            List <int> B = new List<int>();
            int a = 0;
            foreach (String item in PINTURAS)
            {
                if (item ==";;;;")
                {
                    B.Add(a);
                }
                a++;
            }
            foreach (int item in B)
            {
                PINTURAS.RemoveAt(item);
            }


            return PINTURAS;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            TSM.Model m = new Model();
            if (m.GetProjectInfo().ProjectNumber == label10.Text)
            {
                if (comboBox3.Text == "SIM")
                {
                    List<TSM.ModelObject> ojectos = new List<TSM.ModelObject>();
                    ArrayList filtros = new ArrayList();
                    BinaryFilterExpression filterExpression;
                    Random r = new Random();
                    foreach (var item in listBox1.Items)
                    {
                        double n1 = r.Next(1, 10) * 0.1;
                        double n2 = r.Next(1, 10) * 0.1;
                        double n3 = r.Next(1, 10) * 0.1;
                        TSM.UI.Color cor = new TSM.UI.Color(n1, n2, n3);
                        filterExpression = FAZFILTROS("lote_number", item.ToString());
                        TSM.ModelObjectEnumerator modelObjectSelector = m.GetModelObjectSelector().GetObjectsByFilter(filterExpression);
                        modelObjectSelector.SelectInstances = false;
                        while (modelObjectSelector.MoveNext())
                        {
                            TSM.Part a = modelObjectSelector.Current as TSM.Part;
                            if (a != null)
                            {
                                ojectos.Add(a);
                            }
                        }

                        ModelObjectVisualization.SetTemporaryState(ojectos, cor);
                        ojectos.Clear();
                    }
                }
                else
                {
                    TeklaStructures.Connect();
                    TeklaStructures.ExecuteScript("akit.Callback(\"acmd_redraw_selected_view\", \"\", \"main_frame\");");
                    TeklaStructures.Disconnect();
                }
            }
        }

        BinaryFilterExpression FAZFILTROS(string atributo, string parametro)
        {


            BinaryFilterExpression filterExpression = new BinaryFilterExpression(new TemplateFilterExpressions.CustomString(atributo), StringOperatorType.IS_EQUAL, new StringConstantFilterExpression(parametro));

            return filterExpression;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            TSM.Model Model = new TSM.Model();
            if (Model.GetProjectInfo().ProjectNumber == label10.Text)
            {
                ArrayList PECAS = SELECTMODEL();
                foreach (TSM.Part item in PECAS)
                {
                    item.SetUserProperty("lote_number", "");
                    item.SetUserProperty("lote_data", "");
                    item.SetUserProperty("Local_descarga", "");

                }
            }
        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void comboBox4_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox7.Text == "S\\PINTURA")
            {
                comboBox6.Items.Clear();
                comboBox6.Text = "";
            }
            else
            {

                comboBox6.Items.Add("DECAPADO E PINTADO");
                comboBox6.Items.Add("PINTADO");
                comboBox6.Items.Add("DECAPADO");
                comboBox6.Items.Add("GALVANIZADO");
                comboBox6.Items.Add("GALVANIZADO E PINTADO");
                comboBox6.Items.Add("METALIZADO E PINTADO");
                comboBox6.Items.Add("METALIZADO");
                comboBox6.Items.Add("LACADO");
                comboBox6.Items.Add("OUTRO");
                comboBox6.SelectedIndex = 0;
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
                EnviaproPriedadeConj(CONJUNTOS_DA_SELEÇAO(), "pintura", comboBox7.Text);
                EnviaproPriedadeConj(CONJUNTOS_DA_SELEÇAO(), "OperacaoFabrica", comboBox6.Text);
        }

        internal static void EnviaproPriedadeConj(ArrayList Assemblys, string propriedade, string valor)
        {
            foreach (TSM.Assembly assembly in Assemblys)
            {
                assembly.SetUserProperty(propriedade, valor);
            }
        }//envia para o tekla recebe uma lista de conjuntos o campo do objects.inp a preencher e o valor que quero prencher//
        internal ArrayList CONJUNTOS_DA_SELEÇAO()
        {
            lblinfo.Text = "Comunicar com tekla";
            ArrayList segParts = new ArrayList();
            ArrayList assembly = new ArrayList();
            int a = 0;

            TSM.Model model = new TSM.Model();
            TSM.ModelObjectEnumerator modelEnumerator = new TSM.UI.ModelObjectSelector().GetSelectedObjects();

            while (modelEnumerator.MoveNext())
            {
                TSM.Part part = modelEnumerator.Current as TSM.Part;
                TSM.Assembly ass = modelEnumerator.Current as TSM.Assembly;
                if (ass==null&& part!=null)
                {
                    ass = part.GetAssembly();
                }

                if (ass != null)
                {
                    if (!segParts.Contains(ass))
                    {
                        assembly.Add(ass);
                        a++;
                        lblinfo.Text = "procurar conjuntos " + a + " encontrados";
                    }
                }
            }

            lblinfo.Text = "" + assembly.Count + " Conjuntos encontrados";

            return assembly;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            TSM.Model Model = new TSM.Model();
            if (Model.GetProjectInfo().ProjectNumber == label10.Text)
            {
                ArrayList PECAS = SELECTMODEL();
                foreach (TSM.Part item in PECAS)
                {
                    item.SetUserProperty("lote_number", "");
                    item.SetUserProperty("lote_data", "");
                    item.SetUserProperty("Local_descarga", "");
                    item.GetAssembly().SetUserProperty("pintura", "");
                    item.GetAssembly().SetUserProperty("OperacaoFabrica", "");
                }
                lblinfo.Text = "" + PECAS.Count + " PARÂMETROS DE LOTE E ESQUEMA ELIMINADOS";
            }
            
        }
    }
}