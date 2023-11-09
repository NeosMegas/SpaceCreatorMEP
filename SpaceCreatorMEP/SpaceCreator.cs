using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SpaceCreatorMEP
{
    internal class SpaceCreator
    {
        UIDocument doc = null;

        public SpaceCreator(UIDocument document)
        {
            this.doc = document;
        }

        public static ICollection<ElementId> WarnElements = new List<ElementId>();

        #region Space Creation macros
        #region Classes
        public class RoomsData
        {
            public Level RoomsLevel { get; set; }
            public Level UpperRoomLevel { get; set; }
            public List<Room> RoomsList { get; set; }
            public RoomsData()
            {

            }
            public RoomsData(Level level, Level upperRoomLevel, List<Room> roomsList)
            {
                RoomsLevel = level;
                RoomsList = roomsList;
                UpperRoomLevel = upperRoomLevel;
            }
            public Level GetUpperLevel()
            {
                return UpperRoomLevel;
            }
        }
        public class SpacesData
        {
            public Level SpacesLevel { get; set; }
            public List<Space> SpacesList { get; set; }
            public SpacesData()
            {

            }
            public SpacesData(Level level, List<Space> spacesList)
            {
                SpacesLevel = level;
                SpacesList = spacesList;
            }
        }
        public class LevelsData
        {
            public string LevelName { get; set; }
            public double LevelElevation { get; set; }
            public int ElementsCount { get; set; }
            public LevelsData()
            {

            }
            public LevelsData(string levelName, double levelElevation, int elementCount)
            {
                LevelName = levelName;
                LevelElevation = levelElevation;
                ElementsCount = elementCount;
            }
        }
        #endregion
        #region Methods for Classes
        private List<RoomsData> GetRooms(Document linkedDoc)
        {
            List<RoomsData> arRooms = new List<RoomsData>();
            List<Level> arLevels = GetLevels(linkedDoc).OrderBy(l => l.Elevation).ToList();
            for (int i = 0; i < arLevels.Count; i++)
            {
                List<Room> roomsInLevel = GetRoomsByLevel(linkedDoc, arLevels[i]);
                if (roomsInLevel.Count > 0)
                {
                    Level upperLevel = arLevels[i];
                    int next_level = i + 1;
                    while ((next_level < arLevels.Count) && (GetRoomsByLevel(linkedDoc, arLevels[next_level]).Count == 0))
                    {
                        next_level++;
                    }
                    if (next_level < arLevels.Count)
                    {
                        upperLevel = arLevels[next_level];
                    }

                    arRooms.Add(new RoomsData(arLevels[i], upperLevel, roomsInLevel));
                }
            }
            return arRooms;
        }
        private List<SpacesData> GetSpaces(Document currentDoc)
        {
            List<SpacesData> mepSpaces = new List<SpacesData>();
            List<Level> curLevels = GetLevels(currentDoc);
            foreach (Level curLevel in curLevels)
            {
                List<Space> spacesInLevel = GetSpacesByLevel(currentDoc, curLevel);
                if (spacesInLevel.Count > 0)
                {
                    mepSpaces.Add(new SpacesData(curLevel, spacesInLevel));
                }
            }
            return mepSpaces;
        }
        private List<Room> GetRoomsByLevel(Document _doc, Level _level)
        {
            return new FilteredElementCollector(_doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_Rooms)
                .Cast<Room>()
                .Where(e => (e as SpatialElement).Area > 0 && e.Level.Id.IntegerValue.Equals(_level.Id.IntegerValue))
                .ToList();
        }
        private List<Space> GetSpacesByLevel(Document _doc, Level _level)
        {
            return new FilteredElementCollector(_doc)
                .WhereElementIsNotElementType()
                .OfCategory(BuiltInCategory.OST_MEPSpaces)
                .Cast<Space>()
                .Where(e => e.Level.Id.IntegerValue.Equals(_level.Id.IntegerValue) && e.Volume != 0)
                .ToList();
        }
        private List<Level> GetLevels(Document _doc)
        {
            return new FilteredElementCollector(_doc)
                .OfCategory(BuiltInCategory.OST_Levels)
                .WhereElementIsNotElementType()
                .Cast<Level>()
                .ToList();
        }
        private List<LevelsData> MatchLevels(List<Level> linkedLevelList)
        {
            List<LevelsData> levelsData = new List<LevelsData>();
            foreach (Level checkedLevel in linkedLevelList)
            {
                if (GetRoomsByLevel(linkedDOC, checkedLevel).Count() > 0)
                {
                    if (null == GetLevelByElevation(DOC, checkedLevel.Elevation))
                    {
                        levelsData.Add(new LevelsData(checkedLevel.Name, checkedLevel.Elevation, GetRoomsByLevel(linkedDOC, checkedLevel).Count()));
                    }
                }
            }
            return levelsData;
        }
        private bool CreateLevels(List<LevelsData> elevList)
        {
            bool res = false;
            using (Transaction trLevels = new Transaction(DOC, "Создание уровней"))
            {
                trLevels.Start();
                foreach (LevelsData lData in elevList)
                {
                    using (SubTransaction sLevel = new SubTransaction(DOC))
                    {
                        sLevel.Start();
                        Level newLevel = Level.Create(DOC, lData.LevelElevation);
                        newLevel.Name = "АР_" + lData.LevelName;//"АР_"+UnitUtils.ConvertFromInternalUnits(elevation, DisplayUnitType.DUT_MILLIMETERS);
                        sLevel.Commit();
                        res = true;
                    }
                }
                trLevels.Commit();
            }
            return res;
        }
        #endregion
        #region SelectionFIlter HelperClass and Method for selecting RvtLinkInstances
        public class RvtLinkInstanceFilter : ISelectionFilter
        {
            public bool AllowElement(Element element)
            {
                if (element is RevitLinkInstance)
                {
                    return true;
                }
                return false;
            }

            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }
        private Reference GetARLinkReference()
        {
            Selection arSelection = doc.Selection;
            try
            {
                return arSelection.PickObject(ObjectType.Element, new RvtLinkInstanceFilter(), "Выберите экземпляр размещенной связи АР");
            }
            catch (Exception)
            {
                //user abort selection or other
                return null;
            }
        }
        #endregion
        #region FailureProcessor HelperClass for Spaces
        public class SpaceExistWarner : IFailuresPreprocessor
        {
            public FailureProcessingResult PreprocessFailures(FailuresAccessor a)
            {
                IList<FailureMessageAccessor> failures = a.GetFailureMessages();
                foreach (FailureMessageAccessor f in failures)
                {
                    FailureDefinitionId id = f.GetFailureDefinitionId();
                    if (BuiltInFailures.GeneralFailures.DuplicateValue == id)
                    {

                        a.DeleteWarning(f);
                    }
                    if (BuiltInFailures.RoomFailures.RoomsInSameRegionSpaces == id)
                    {
                        WarnElements = f.GetFailingElementIds();
                        a.DeleteWarning(f);
                        return FailureProcessingResult.ProceedWithRollBack;
                    }

                }
                return FailureProcessingResult.Continue;
            }
        }
        #endregion
        #region Helper Method for Levels by Elevation
        private Level GetLevelByElevation(Document _doc, double _elevation)
        {
            return new FilteredElementCollector(_doc)
                .OfClass(typeof(Level))
                .Cast<Level>()
                .Where(l => l.Elevation.Equals(_elevation))
                .FirstOrDefault();
        }
        #endregion
        List<SpacesData> spacesDataList;
        List<RoomsData> roomsDataList;
        Document DOC;
        Document linkedDOC;
        List<LevelsData> levelsDataCreation;
        Transform tGlobal;
        Stopwatch sTimer = new Stopwatch();

        public void CreateSpaces()
        {
            //Check if Document is opened in UI and Document is Project
            sTimer.Reset();
            if (null != doc && !doc.Document.IsFamilyDocument)
            {
                DOC = doc.Document;
                //Start taskdialog for select link
                if (TdSelectARLink() == 0)
                {
                    return;
                }
                Reference arLink = GetARLinkReference();
                if (null != arLink)
                {
                    //Get value for Room bounding in RevitLinkType
                    RevitLinkInstance selInstance = DOC.GetElement(arLink.ElementId) as RevitLinkInstance;
                    RevitLinkType lnkType = DOC.GetElement(selInstance.GetTypeId()) as RevitLinkType;
                    bool boundingWalls = Convert.ToBoolean(lnkType.get_Parameter(BuiltInParameter.WALL_ATTR_ROOM_BOUNDING).AsInteger());
                    //Get coordination transformation for link
                    tGlobal = selInstance.GetTotalTransform();
                    //Get Document from RvtLinkInstance
                    linkedDOC = selInstance.GetLinkDocument();
                    //Check for valid Document and Room bounding value checked
                    if (null != linkedDOC && !boundingWalls)
                    {
                        TaskDialog.Show("Ошибка", "Нет загруженной связи АР или в связанном файле не включен поиск границ помещения!\nДля размещения пространств необходимо включить этот параметр");
                        return;
                    }
                    //Mainline code
                    //Get placed Spaces and Levels information
                    spacesDataList = GetSpaces(DOC);
                    roomsDataList = GetRooms(linkedDOC);
                    //Check if Spaces placed
                    if (roomsDataList.Count > 0)
                    {
                        if (spacesDataList.Count == 0)
                        {
                            switch (TdFirstTime())
                            {
                                case 0:
                                    return;
                                case 1:
                                    AnaliseAR();
                                    break;
                                default:
                                    return;
                            }
                        }
                        else
                        {
                            AnaliseAR();
                        }
                    }
                    else
                    {
                        TaskDialog tDialog = new TaskDialog("No Rooms in link");
                        tDialog.Title = "Нет помещений";
                        tDialog.MainInstruction = "В выбранном экземпляре связи нет помещений.";
                        tDialog.MainIcon = TaskDialogIcon.TaskDialogIconWarning;
                        tDialog.TitleAutoPrefix = false;
                        tDialog.CommonButtons = TaskDialogCommonButtons.Close;
                        tDialog.Show();
                    }

                }
            }
        }

        private void AnaliseAR()
        {
            levelsDataCreation = MatchLevels(roomsDataList.Select(r => r.RoomsLevel).ToList());
            if (levelsDataCreation.Count > 0)
            {
                switch ((TdLevelsCreate()))
                {
                    case 0:
                        return;
                    case 1:
                        sTimer.Start();
                        CreateLevels(levelsDataCreation);
                        sTimer.Stop();
                        break;
                    default:

                        break;
                }
            }
            switch (TdSpacesPlace())
            {
                case 0:
                    return;
                case 1:
                    sTimer.Start();
                    CreateSpByRooms(true);
                    break;
                case 2:
                    sTimer.Start();
                    CreateSpByRooms(false);
                    break;
                default:

                    break;
            }
        }
        #region TaskDialogs
        private int TdSpacesPlace()
        {
            TaskDialog td = new TaskDialog("Spaces place Type");
            td.Id = "ID_TaskDialog_Spaces_Type_Place";
            td.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
            td.Title = "Создание/обновление пространств";
            td.TitleAutoPrefix = false;
            td.AllowCancellation = true;
            // Message related stuffs
            td.MainInstruction = "Настройка задания верхнего предела и смещения для размещения пространств";
            td.MainContent = "Выберите способ создания/обновления пространств";
            td.ExpandedContent = "При выборе варианта по помещениям - пространства создаются с копированием всех настроект для привязки верхнего уровня и смещения из помещений. Так же будут дополнительно скопированы уровни для верхнего предела, если таковые отсутсвуют.\n" +
                "При выборе варианта по уровням - верхним пределом для пространств будет следующий уровень с помещениями и значение смещения 0 мм. Если такового не существует (последний уровень с помещениями) используется смещение 3500 мм.";
            // Command link stuffs
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "По помещениям");
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink2, "По уровням");
            // Dialog showup stuffs
            TaskDialogResult tdRes = td.Show();
            switch (tdRes)
            {
                case TaskDialogResult.Cancel:
                    return 0;
                case TaskDialogResult.CommandLink1:
                    return 1;
                case TaskDialogResult.CommandLink2:
                    return 2;
                default:
                    throw new Exception("Invalid value for TaskDialogResult");
            }
        }
        private int TdLevelsCreate()
        {
            TaskDialog td = new TaskDialog("Levels Create for Spaces");

            td.Id = "ID_TaskDialog_Create_Levels";
            td.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
            td.Title = "Создание недостающих уровней";
            td.TitleAutoPrefix = false;
            td.AllowCancellation = true;
            // Message related stuffs
            td.MainInstruction = "В текущем проекте не хватет уровней для создания инженерных пространств";
            td.MainContent = "Создать уровни в текущем проекте";
            td.ExpandedContent = "В выбранном файле архитектуры имеются помещения, размещенные на уровнях, отсутсвующих в текущем проекте.\nНеобходимо создать уровни в текущем проекте для автоматического размещения инженерных пространств\n" +
                "Будут созданы уровни с префиксом АР_имя уровня для всех помещений";
            // Command link stuffs
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Продолжить");
            // Dialog showup stuffs
            TaskDialogResult tdRes = td.Show();
            switch (tdRes)
            {
                case TaskDialogResult.Cancel:
                    return 0;
                case TaskDialogResult.CommandLink1:
                    return 1;
                //break;
                default:
                    throw new Exception("Invalid value for TaskDialogResult");
            }
        }
        private int TdFirstTime()
        {
            TaskDialog td = new TaskDialog("First time Spaces placing");
            td.Id = "ID_TaskDialog_Place_Spaces_Type";
            td.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
            td.Title = "Размещение инженерных пространств";
            td.TitleAutoPrefix = false;
            td.AllowCancellation = true;
            // Message related stuffs
            td.MainInstruction = "Втекущем проекте нет размещенных инженерных пространств";
            td.MainContent = "Проанализировать связанный файл на наличие помещений";
            // Command link stuffs
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Продолжить");
            // Dialog showup stuffs
            TaskDialogResult tdRes = td.Show();
            switch (tdRes)
            {
                case TaskDialogResult.Cancel:
                    return 0;
                case TaskDialogResult.CommandLink1:
                    return 1;
                //break;
                default:
                    throw new Exception("Invalid value for TaskDialogResult");
            }
        }
        private int TdSelectARLink()
        {
            TaskDialog td = new TaskDialog("Select Link File");

            td.Id = "ID_TaskDialog_Select_AR";
            td.MainIcon = TaskDialogIcon.TaskDialogIconInformation;
            td.Title = "Выбор экземпляра размещенной связи";
            td.TitleAutoPrefix = false;
            td.AllowCancellation = true;
            // Message related stuffs
            td.MainInstruction = "Выберите экземпляр размещенной связи для поиска помещений";
            // Command link stuffs
            td.AddCommandLink(TaskDialogCommandLinkId.CommandLink1, "Выбрать");
            // Dialog showup stuffs
            TaskDialogResult tdRes = td.Show();
            switch (tdRes)
            {
                case TaskDialogResult.Cancel:
                    return 0;
                case TaskDialogResult.CommandLink1:
                    return 1;
                default:
                    throw new Exception("Invalid value for TaskDialogResult");
            }
        }
        #endregion

        private void CreateSpByRooms(bool RoomLimits)
        {
            int levels = 0;
            int sCreated = 0;
            int sUpdated = 0;
            double defLimitOffset = 3500;
            using (TransactionGroup crTrans = new TransactionGroup(DOC, "Создание пространств"))
            {
                crTrans.Start();
                foreach (RoomsData roomsData in roomsDataList)
                {
                    Level lLevel = roomsData.RoomsLevel;
                    Level localLevel = GetLevelByElevation(DOC, lLevel.Elevation);

                    if (null != localLevel)
                    {
                        levels++;
                        foreach (Room lRoom in roomsData.RoomsList)
                        {
                            string RoomName = lRoom.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                            string RoomNumber = lRoom.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString();
                            if (lRoom.Level.Elevation.Equals(localLevel.Elevation))
                            {
                                using (Transaction tr = new Transaction(DOC, "Create space"))
                                {
                                    tr.Start();
                                    FailureHandlingOptions failOpt = tr.GetFailureHandlingOptions();
                                    failOpt.SetFailuresPreprocessor(new SpaceExistWarner());
                                    tr.SetFailureHandlingOptions(failOpt);
                                    LocationPoint lp = lRoom.Location as LocationPoint;
                                    XYZ rCoord = tGlobal.OfPoint(lp.Point);
                                    Space sp = null;
                                    UV spLocPoint = new UV(rCoord.X, rCoord.Y);
                                    sp = DOC.Create.NewSpace(localLevel, spLocPoint);
                                    TransactionStatus trStat = tr.Commit(failOpt);
                                    //Space not exists, change name and number for new space
                                    if (trStat == TransactionStatus.Committed)
                                    {
                                        tr.Start();
                                        sp.get_Parameter(BuiltInParameter.ROOM_NAME).Set(RoomName);
                                        sp.get_Parameter(BuiltInParameter.ROOM_NUMBER).Set(RoomNumber);
                                        //TryToRenameSpace(tr,sp,RoomName,RoomNumber);
                                        if (RoomLimits)
                                        {
                                            Level upperLevel = GetLevelByElevation(DOC, lRoom.UpperLimit.Elevation);
                                            if (null == upperLevel)
                                            {
                                                upperLevel = Level.Create(DOC, lRoom.UpperLimit.Elevation);
                                                upperLevel.Name = "АР_" + lRoom.UpperLimit.Name;
                                            }
                                        }
                                        else
                                        {
                                            if (roomsData.RoomsLevel == roomsData.UpperRoomLevel)
                                            {
                                                Level upperLevel = localLevel;
                                                //sp.get_Parameter(BuiltInParameter.ROOM_UPPER_LEVEL).Set(localLevel.LevelId);
                                                sp.UpperLimit = GetLevelByElevation(DOC, localLevel.Elevation);
                                                sp.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(UnitUtils.ConvertToInternalUnits(defLimitOffset, UnitTypeId.Millimeters));
                                            }
                                            else
                                            {
                                                Level upperLevel = roomsData.UpperRoomLevel;
                                                sp.UpperLimit = GetLevelByElevation(DOC, upperLevel.Elevation);
                                                sp.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(0);
                                            }
                                        }
                                        tr.Commit(failOpt);
                                        sCreated++;
                                    }
                                    //If Space placed in same area. Transaction creating space Rolledback
                                    else
                                    {
                                        foreach (ElementId eId in WarnElements)
                                        {
                                            if (null != DOC.GetElement(eId) && DOC.GetElement(eId) is Space)
                                            {
                                                Space wSpace = DOC.GetElement(eId) as Space;
                                                //bool updated = false;
                                                string SpaceName = wSpace.get_Parameter(BuiltInParameter.ROOM_NAME).AsString();
                                                string SpaceNumber = wSpace.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString();
                                                LocationPoint sLp = wSpace.Location as LocationPoint;

                                                tr.Start();
                                                sLp.Point = rCoord;
                                                wSpace.get_Parameter(BuiltInParameter.ROOM_NAME).Set(RoomName);
                                                wSpace.get_Parameter(BuiltInParameter.ROOM_NUMBER).Set(RoomNumber);
                                                if (RoomLimits)
                                                {
                                                    Level upperLevel = GetLevelByElevation(DOC, lRoom.UpperLimit.Elevation);
                                                    if (null == upperLevel)
                                                    {
                                                        upperLevel = Level.Create(DOC, lRoom.UpperLimit.Elevation);
                                                        upperLevel.Name = "АР_" + lRoom.UpperLimit.Name;
                                                    }
                                                    wSpace.UpperLimit = upperLevel;
                                                    wSpace.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(lRoom.LimitOffset);
                                                }
                                                else
                                                {
                                                    if (roomsData.RoomsLevel == roomsData.UpperRoomLevel)
                                                    {
                                                        Level upperLevel = localLevel;
                                                        wSpace.UpperLimit = upperLevel;
                                                        wSpace.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(UnitUtils.ConvertToInternalUnits(defLimitOffset, UnitTypeId.Millimeters));
                                                    }
                                                    else
                                                    {
                                                        Level upperLevel = roomsData.UpperRoomLevel;
                                                        wSpace.UpperLimit = GetLevelByElevation(DOC, upperLevel.Elevation);
                                                        wSpace.get_Parameter(BuiltInParameter.ROOM_UPPER_OFFSET).Set(0);
                                                    }
                                                }
                                                tr.Commit();
                                                sUpdated++;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                crTrans.Assimilate();
            }
            sTimer.Stop();
            TimeSpan resTimer = sTimer.Elapsed;
            TaskDialog tDialog = new TaskDialog("Rsult data");
            tDialog.Title = "Отчёт";
            tDialog.MainInstruction = String.Format("Обновлено {0}\nСоздано {1}", sUpdated, sCreated);
            tDialog.MainIcon = TaskDialogIcon.TaskDialogIconShield;
            tDialog.TitleAutoPrefix = false;
            tDialog.FooterText = String.Format("Общее время {0:D2}:{1:D2}:{2:D3}", resTimer.Minutes, resTimer.Seconds, resTimer.Milliseconds);
            tDialog.CommonButtons = TaskDialogCommonButtons.Close;
            tDialog.Show();
        }
        #endregion

    }
}
