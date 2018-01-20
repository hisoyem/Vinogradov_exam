using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Inventor;

namespace Vinogradov_exam
{
    public partial class Form1 : Form
    {
        /// ThisApplication - Объект для определения активного состояния Инвентора
        /// </summary>
        private Inventor.Application ThisApplication = null;
        /// <summary>
        /// Словарь для хранения ссылок на документы деталей
        /// </summary>
        private Dictionary<string, PartDocument> oPartDoc = new Dictionary<string, PartDocument>();
        /// <summary>
        /// Словарь для хранения ссылок на определения деталей
        /// </summary>
        private Dictionary<string, PartComponentDefinition> oCompDef = new
        Dictionary<string, PartComponentDefinition>();
        /// <summary>
        /// Словарь для хранения ссылок на инструменты создания деталей
        /// </summary>
        private Dictionary<string, TransientGeometry> oTransGeom = new
        Dictionary<string, TransientGeometry>();
        /// <summary>
        /// Словарь для хранения ссылок на транзакции редактирования
        /// </summary>
        private Dictionary<string, Transaction> oTrans = new Dictionary<string, Transaction>();
        /// <summary>
        /// Словарь для хранения имен сохраненных документов деталей
        /// </summary>
        private Dictionary<string, string> oFileName = new Dictionary<string, string>();

        public static double A, B, C, D, E, F, G, H, I, J, K, L, M = 0;

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("C:/Program Files/Autodesk/Inventor2017/Bin/Inventor.exe");
                label1.Text = "Inventor запущен.";
            }
            catch
            {
                System.Diagnostics.Process.Start("C:/Program Files/Autodesk/Inventor 2017/Bin/Inventor.exe");
                label1.Text = "Inventor запущен.";
            }
        }
        public Form1()
        {
            InitializeComponent();
            try
            {
                ThisApplication = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                if (ThisApplication != null) label1.Text = "Inventor запущен.";
            }
            catch
            {
                label1.Text = "Запустите Inventor!";
            }
        }
        private void New_document_Name(string Name)
        {
            // Новый документ детали
            oPartDoc[Name] = (PartDocument)ThisApplication.Documents.Add(DocumentTypeEnum.kPartDocumentObject, ThisApplication.FileManager.GetTemplateFile(DocumentTypeEnum.kPartDocumentObject));
            // Новое определение
            oCompDef[Name] = oPartDoc[Name].ComponentDefinition;
            // Выбор инструментов
            oTransGeom[Name] = ThisApplication.TransientGeometry;
            // Создание транзакции
            oTrans[Name] = ThisApplication.TransactionManager.StartTransaction(
            ThisApplication.ActiveDocument, "Create Sample");
            // Имя файла
            oFileName[Name] = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                ThisApplication = (Inventor.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Inventor.Application");
                if (ThisApplication != null) label1.Text = "Inventor запущен.";
            }
            catch
            {
                MessageBox.Show("Запустите Inventor!");
                return;
            }

            A = Convert.ToDouble(textBox1.Text);
            B = Convert.ToDouble(textBox2.Text);
            C = Convert.ToDouble(textBox3.Text);
            D = Convert.ToDouble(textBox4.Text);
            E = Convert.ToDouble(textBox5.Text);
            F = Convert.ToDouble(textBox6.Text);
            G = Convert.ToDouble(textBox7.Text);
            H = Convert.ToDouble(textBox8.Text);
            I = Convert.ToDouble(textBox9.Text);
            J = Convert.ToDouble(textBox10.Text);
            K = Convert.ToDouble(textBox11.Text);
            L = Convert.ToDouble(textBox12.Text);
            M = Convert.ToDouble(textBox13.Text);

            //Перевод мм в см
            A /= 10; B /= 10; C /= 10; D /= 10; E /= 10; F /= 10; G /= 10; H /= 10; I /= 10; J /= 10; K /= 10; L /= 10; M /= 10;

            SketchPoint[] point = new SketchPoint[101];
            SketchLine[] lines = new SketchLine[101];
            SketchArc[] arc = new SketchArc[101];

            //Эскиз в осях XY
            New_document_Name("Деталь");
            oPartDoc["Деталь"].DisplayName = "Деталь";
            PlanarSketch oSketch = oCompDef["Деталь"].Sketches.Add(oCompDef["Деталь"].WorkPlanes[3]);

            //Построение эскиза
            point[0] = oSketch.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(0, 0));
            point[1] = oSketch.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A, 0));
            lines[0] = oSketch.SketchLines.AddByTwoPoints(point[0], point[1]);

            point[2] = oSketch.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A, D));
            lines[1] = oSketch.SketchLines.AddByTwoPoints(point[1], point[2]);

            point[3] = oSketch.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d((A-E)/2+E, D));
            lines[2] = oSketch.SketchLines.AddByTwoPoints(point[2], point[3]);

            point[4] = oSketch.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d((A-E)/2, D));
            point[5] = oSketch.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(0, D));
            lines[3] = oSketch.SketchLines.AddByTwoPoints(point[4], point[5]);
            lines[4] = oSketch.SketchLines.AddByTwoPoints(point[0], point[5]);

            point[6] = oSketch.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d((A - E) / 2 + E, D - I));
            point[7] = oSketch.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d((A - E) / 2, D - I));
            lines[5] = oSketch.SketchLines.AddByTwoPoints(point[4], point[7]);
            lines[6] = oSketch.SketchLines.AddByTwoPoints(point[3], point[6]);
            lines[7] = oSketch.SketchLines.AddByTwoPoints(point[6], point[7]);

            //Конец эскиза
            oTrans["Деталь"].End();

            //Выдавливание первого эскиза (прямоугольника с выемкой)
            Profile oProfile = default(Profile);
            oProfile = (Profile)oSketch.Profiles.AddForSolid();
            ExtrudeFeature oExtrude;
            oExtrude = oCompDef["Деталь"].Features.ExtrudeFeatures.AddByDistanceExtent(oProfile, B, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, PartFeatureOperationEnum.kJoinOperation);

            //Верхний эскиз для выдавливания "рампы"
            PlanarSketch oSketch1 = oCompDef["Деталь"].Sketches.Add(oCompDef["Деталь"].WorkPlanes[3]);
            Transaction oTrans1 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample");

            //Построение точек
            point[8] = oSketch1.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(0, D));
            point[9] = oSketch1.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A, D));
            lines[8] = oSketch1.SketchLines.AddByTwoPoints(point[8], point[9]);

            point[10] = oSketch1.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d((A - C) / 2, D + J));
            point[11] = oSketch1.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A - (A - C) / 2, D + J));
            point[12] = oSketch1.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d((A - 2 * K) / 2, D + J));
            point[13] = oSketch1.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d((A - 2 * K) / 2 + 2*K, D + J));
            lines[9] = oSketch1.SketchLines.AddByTwoPoints(point[11], point[13]);
            lines[10] = oSketch1.SketchLines.AddByTwoPoints(point[10], point[12]);
            lines[11] = oSketch1.SketchLines.AddByTwoPoints(point[8], point[10]);
            lines[12] = oSketch1.SketchLines.AddByTwoPoints(point[9], point[11]);

            //Арка
            point[11] = oSketch1.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A / 2, D + J ));
            arc[0] = oSketch1.SketchArcs.AddByCenterStartEndPoint(point[11], point[12], point[13]);

            //Конец эскиза
            oTrans1.End();

            //Выдавливание "рампы"
            Profile oProfile1 = default(Profile);
            oProfile1 = (Profile)oSketch1.Profiles.AddForSolid();
            ExtrudeFeature oExtrude1;
            oExtrude1 = oCompDef["Деталь"].Features.ExtrudeFeatures.AddByDistanceExtent(oProfile1, M, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, PartFeatureOperationEnum.kJoinOperation);

            //Эскиз вырезаний при выдавливании.
            PlanarSketch oSketch2 = oCompDef["Деталь"].Sketches.Add(oCompDef["Деталь"].WorkPlanes[3]);
            Transaction oTrans2 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample");

            //Построение эскиза!
            point[13] = oSketch2.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(0,0));
            point[14] = oSketch2.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A, 0));
            lines[13] = oSketch2.SketchLines.AddByTwoPoints(point[13], point[14]);

            point[15] = oSketch2.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(0, G));
            lines[14] = oSketch2.SketchLines.AddByTwoPoints(point[13], point[15]);

            point[16] = oSketch2.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A, G));
            lines[15] = oSketch2.SketchLines.AddByTwoPoints(point[15], point[16]);
            lines[16] = oSketch2.SketchLines.AddByTwoPoints(point[16], point[14]);

            //Конец эскиза
            oTrans2.End();

            //Выдавливание (вычетание) прямоугольника снизу
            Profile oProfile2 = default(Profile);
            oProfile2 = (Profile)oSketch2.Profiles.AddForSolid();
            ExtrudeFeature oExtrude2;
            oExtrude2 = oCompDef["Деталь"].Features.ExtrudeFeatures.AddByDistanceExtent(oProfile2, F, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, PartFeatureOperationEnum.kCutOperation);

            //New Sketch, где последний эскиз
            PlanarSketch oSketch3 = oCompDef["Деталь"].Sketches.Add(oCompDef["Деталь"].WorkPlanes[3]);
            Transaction oTrans3 = ThisApplication.TransactionManager.StartTransaction(ThisApplication.ActiveDocument, "Create Sample");

            //Построение эскиза
            point[17] = oSketch3.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(0, D + J));
            point[18] = oSketch3.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A, D + J));
            lines[16] = oSketch3.SketchLines.AddByTwoPoints(point[17], point[18]);

            point[19] = oSketch3.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(0, D + J - H));
            lines[17] = oSketch3.SketchLines.AddByTwoPoints(point[17], point[19]);

            point[20] = oSketch3.SketchPoints.Add(oTransGeom["Деталь"].CreatePoint2d(A, D + J - H));
            lines[18] = oSketch3.SketchLines.AddByTwoPoints(point[20], point[18]);
            lines[19] = oSketch3.SketchLines.AddByTwoPoints(point[20], point[19]);

            //Закрытие эскиза
            oTrans3.End();

            //Выдавливание (отрицательное) так называемой "рампы" ниже
            Profile oProfile3 = default(Profile);
            oProfile3 = (Profile)oSketch3.Profiles.AddForSolid();
            ExtrudeFeature oExtrude3;
            oExtrude3 = oCompDef["Деталь"].Features.ExtrudeFeatures.AddByDistanceExtent(oProfile3, L, PartFeatureExtentDirectionEnum.kSymmetricExtentDirection, PartFeatureOperationEnum.kCutOperation);

        }
    }
}
