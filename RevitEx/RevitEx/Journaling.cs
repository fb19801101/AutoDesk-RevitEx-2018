using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using RevitDB = Autodesk.Revit.DB;
using RevitUI = Autodesk.Revit.UI;
using RevitCreation = Autodesk.Revit.Creation;
using RevitAttributes = Autodesk.Revit.Attributes;
using RevitApplicationServices = Autodesk.Revit.ApplicationServices;

namespace RevitEx
{
    public class Journaling
    {
        class WallTypeComparer : IComparer<RevitDB.WallType>
        {
            int IComparer<RevitDB.WallType>.Compare(RevitDB.WallType first, RevitDB.WallType second)
            {
                return string.Compare(first.Name, second.Name);
            }
        }

        RevitUI.ExternalCommandData m_commandData;
        bool m_canReadData;

        RevitDB.XYZ m_startPoint;
        RevitDB.XYZ m_endPoint;
        RevitDB.Level m_createlevel;
        RevitDB.WallType m_createType;

        List<RevitDB.Level> m_levelList;
        List<RevitDB.WallType> m_wallTypeList;

        public ReadOnlyCollection<RevitDB.Level> Levels
        {
            get
            {
                return new ReadOnlyCollection<RevitDB.Level>(m_levelList);
            }
        }

        public ReadOnlyCollection<RevitDB.WallType> WallTypes
        {
            get
            {
                return new ReadOnlyCollection<RevitDB.WallType>(m_wallTypeList);
            }
        }

        public Journaling(RevitUI.ExternalCommandData commandData)
        {
            m_commandData = commandData;
            m_canReadData = (0 < commandData.JournalData.Count) ? true : false;

            m_levelList = new List<RevitDB.Level>();
            m_wallTypeList = new List<RevitDB.WallType>();
            InitializeListData();
        }

        public void Run()
        {
            if (m_canReadData)
            {
                ReadJournalData();
                CreateWall();
            }
            else
            {
                if (!DisplayUI())
                {
                    return;
                }

                CreateWall();
                WriteJournalData();
            }
        }

        public void SetNecessaryData(RevitDB.XYZ startPoint, RevitDB.XYZ endPoint, RevitDB.Level level, RevitDB.WallType type)
        {
            m_startPoint = startPoint;
            m_endPoint = endPoint;
            m_createlevel = level;
            m_createType = type;
        }

        private void InitializeListData()
        {
            if (null == m_wallTypeList || null == m_levelList)
            {
                throw new Exception("necessary data members don't initialize.");
            }

            RevitDB.Document document = m_commandData.Application.ActiveUIDocument.Document;
            RevitDB.FilteredElementCollector filteredElementCollector = new RevitDB.FilteredElementCollector(document);
            filteredElementCollector.OfClass(typeof(RevitDB.WallType));
            m_wallTypeList = filteredElementCollector.Cast<RevitDB.WallType>().ToList<RevitDB.WallType>();

            WallTypeComparer comparer = new WallTypeComparer();
            m_wallTypeList.Sort(comparer);

            RevitDB.FilteredElementIterator iter = (new RevitDB.FilteredElementCollector(document)).OfClass(typeof(RevitDB.Level)).GetElementIterator();
            iter.Reset();
            while (iter.MoveNext())
            {
                RevitDB.Level level = iter.Current as RevitDB.Level;
                if (null == level)
                {
                    continue;
                }
                m_levelList.Add(level);
            }
        }

        private void ReadJournalData()
        {
            RevitDB.Document doc = m_commandData.Application.ActiveUIDocument.Document;
            IDictionary<string, string> dataMap = m_commandData.JournalData;
            string dataValue = null;
            dataValue = GetSpecialData(dataMap, "Wall Type Name");
            foreach (RevitDB.WallType type in m_wallTypeList)
            {
                if (dataValue == type.Name)
                {
                    m_createType = type;
                    break;
                }
            }
            if (null == m_createType)
            {
                throw new InvalidDataException("Can't find the wall type from the journal.");
            }

            dataValue = GetSpecialData(dataMap, "Level Id");
            RevitDB.ElementId id = new RevitDB.ElementId(Convert.ToInt32(dataValue));

            m_createlevel = doc.GetElement(id) as RevitDB.Level;
            if (null == m_createlevel)
            {
                throw new InvalidDataException("Can't find the level from the journal.");
            }

            dataValue = GetSpecialData(dataMap, "Start Point");
            m_startPoint = StringToXYZ(dataValue);

            if (m_startPoint.Equals(m_endPoint))
            {
                throw new InvalidDataException("Start point is equal to end point.");
            }
        }

        private bool DisplayUI()
        {
            using (FrmJournaling displayForm = new FrmJournaling(this))
            {
                displayForm.ShowDialog();
                if (DialogResult.OK != displayForm.DialogResult)
                {
                    return false;
                }
            }
            return true;
        }

        private void CreateWall()
        {
            RevitCreation.Application createApp = m_commandData.Application.Application.Create;
            RevitCreation.Document createDoc = m_commandData.Application.ActiveUIDocument.Document.Create;

            RevitDB.Line geometryLine = RevitDB.Line.CreateBound(m_startPoint, m_endPoint);
            if (null == geometryLine)
            {
                throw new Exception("Create the geometry line failed.");
            }

            RevitDB.Wall createdWall = RevitDB.Wall.Create(m_commandData.Application.ActiveUIDocument.Document,
                geometryLine, m_createType.Id, m_createlevel.Id,
                15, m_startPoint.Z + m_createlevel.Elevation, true, true);
            if (null == createdWall)
            {
                throw new Exception("Create the wall failed.");
            }
        }

        private void WriteJournalData()
        {
            IDictionary<string, string> dataMap = m_commandData.JournalData;
            dataMap.Clear();

            dataMap.Add("Wall Type Name", m_createType.Name);
            dataMap.Add("Level Id", m_createlevel.Id.IntegerValue.ToString());
            dataMap.Add("Start Point", XYZToString(m_startPoint));
            dataMap.Add("End Point", XYZToString(m_endPoint));
        }

        private static RevitDB.XYZ StringToXYZ(string pointString)
        {
            double x = 0;
            double y = 0;
            double z = 0;
            string subString;

            subString = pointString.TrimStart('(');
            subString = subString.TrimEnd(')');
            string[] coordinateString = subString.Split(',');
            if (3 != coordinateString.Length)
            {
                throw new InvalidDataException("The point iniformation in journal is incorrect");
            }

            try
            {
                x = Convert.ToDouble(coordinateString[0]);
                y = Convert.ToDouble(coordinateString[1]);
                z = Convert.ToDouble(coordinateString[2]);
            }
            catch (Exception)
            {
                throw new InvalidDataException("The point information in journal is incorrect");

            }

            return new RevitDB.XYZ(x, y, z);
        }

        private static string XYZToString(RevitDB.XYZ point)
        {
            string pointString = "(" + point.X.ToString() + "," + point.Y.ToString() + ","
                + point.Z.ToString() + ")";
            return pointString;
        }

        private static string GetSpecialData(IDictionary<string, string> dataMap, string key)
        {
            string dataValue = dataMap[key];

            if (string.IsNullOrEmpty(dataValue))
            {
                throw new Exception(key + "information is not exits in journal.");
            }
            return dataValue;
        }
    }

    public class AppDocEvents
    {
        private RevitApplicationServices.ControlledApplication m_app;

        public AppDocEvents(RevitApplicationServices.ControlledApplication app)
        {
            m_app = app;
        }

        public void EnableEvents()
        {
            m_app.DocumentClosed += new EventHandler<RevitDB.Events.DocumentClosedEventArgs>(DocumentClosed);
            m_app.DocumentOpened += new EventHandler<RevitDB.Events.DocumentOpenedEventArgs>(DocumentOpened);
            m_app.DocumentSaved += new EventHandler<RevitDB.Events.DocumentSavedEventArgs>(DocumentSaved);
            m_app.DocumentSavedAs += new EventHandler<RevitDB.Events.DocumentSavedAsEventArgs>(DocumentSavedAs);
        }

        public void DisableEvents()
        {
            m_app.DocumentClosed -= new EventHandler<RevitDB.Events.DocumentClosedEventArgs>(DocumentClosed);
            m_app.DocumentOpened -= new EventHandler<RevitDB.Events.DocumentOpenedEventArgs>(DocumentOpened);
            m_app.DocumentSaved -= new EventHandler<RevitDB.Events.DocumentSavedEventArgs>(DocumentSaved);
            m_app.DocumentSavedAs -= new EventHandler<RevitDB.Events.DocumentSavedAsEventArgs>(DocumentSavedAs);
        }

        void DocumentSavedAs(object sender, RevitDB.Events.DocumentSavedAsEventArgs e)
        {
        }

        void DocumentSaved(object sender, RevitDB.Events.DocumentSavedEventArgs e)
        {
        }

        void DocumentOpened(object sender, RevitDB.Events.DocumentOpenedEventArgs e)
        {
        }

        void DocumentClosed(object sender, RevitDB.Events.DocumentClosedEventArgs e)
        {
        }
    }
}