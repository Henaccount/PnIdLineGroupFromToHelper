// Suggested references for this command:
//   acdbmgd.dll
//   acmgd.dll
//   accoremgd.dll
//   PnPProjectManagerMgd.dll
//   PnPDataLinks.dll
//   PnPDataObjects.dll
//   PnIDMgd.dll
//
// Command name: PIDSETFROMTO
//
// Notes:
// - The Pipe Line Group must expose custom properties named from and to
//   (or change RequestedFromProperty / RequestedToProperty below).
// - The >2-equipment workflow uses clickable E1/E2/.../Pick/Cancel options.
//   If the user chooses Pick, they can click a listed equipment object
//   on the active drawing.
// - Direct line/group traversal is database-based and also expands through
//   off-page connector relationships.

using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.ProcessPower.DataLinks;
using Autodesk.ProcessPower.DataObjects;
using Autodesk.ProcessPower.PlantInstance;
using Autodesk.ProcessPower.PnIDObjects;
using Autodesk.ProcessPower.ProjectManager;
using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Linq;
using static System.Runtime.InteropServices.JavaScript.JSType;
using AcadApp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace Plant3D.PnIdTools
{
    public class PidLineGroupFromToCommand
    {
        private const string CommandName = "PIDSETFROMTO";

        private const string RelPipeLineGroup = "PipeLineGroupRelationship";
        private const string RelLineStartAsset = "LineStartAsset";
        private const string RelLineEndAsset = "LineEndAsset";
        private const string RelLineNozzle = "LineNozzle";
        private const string RelLineOffPageConnector = "LineOffPageConnector";
        private const string RelConnectors = "ConnectorsRelationship";
        private const string RelAssetOwnership = "AssetOwnership";

        private const string TableEquipment = "Equipment";
        private const string TableNozzles = "Nozzles";
        private const string TablePnPDataLinks = "PnPDataLinks";
        private const string TablePnPDrawings = "PnPDrawings";

        // Adjust these if your custom property display names differ.
        private const string RequestedFromProperty = "from";
        private const string RequestedToProperty = "to";

        [CommandMethod(CommandName, CommandFlags.Modal)]
        public static void Execute()
        {
            var doc = AcadApp.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                return;
            }

            var ed = doc.Editor;
            var db = doc.Database;

            try
            {
                Project pidProject = GetPidProject();
                if (pidProject == null)
                {
                    ed.WriteMessage("\nNo active P&ID project part was found.");
                    return;
                }

                DataLinksManager dlm = pidProject.DataLinksManager;
                if (dlm == null)
                {
                    ed.WriteMessage("\nThe P&ID DataLinksManager is not available.");
                    return;
                }

                PnPDatabase pnpDb = dlm.GetPnPDatabase();
                if (pnpDb == null)
                {
                    ed.WriteMessage("\nThe P&ID project database is not available.");
                    return;
                }

                ObjectId selectedLineId = PromptForPipeLineSegment(ed);
                if (selectedLineId.IsNull)
                {
                    return;
                }

                int selectedLineRowId = SafeFindRowId(dlm, selectedLineId);
                if (selectedLineRowId <= 0)
                {
                    ed.WriteMessage("\nThe selected Pipe Line Segment is not linked to the P&ID database.");
                    return;
                }

                int lineGroupRowId = GetLineGroupRowId(dlm, selectedLineRowId);
                if (lineGroupRowId <= 0)
                {
                    ed.WriteMessage("\nCould not resolve the Pipe Line Group for the selected segment.");
                    return;
                }

                string lineGroupTag = GetLineGroupDisplayTag(dlm, lineGroupRowId);

                string actualFromPropertyName;
                string actualToPropertyName;
                if (!TryResolveLineGroupPropertyNames(dlm, lineGroupRowId, out actualFromPropertyName, out actualToPropertyName))
                {
                    ed.WriteMessage(
                        "\nThe selected Pipe Line Group does not expose custom properties '" +
                        RequestedFromProperty +
                        "' and '" +
                        RequestedToProperty +
                        "'. Create them in Project Setup or change the constants in the code.");
                    return;
                }

                HashSet<int> lineRowIds = CollectAllRelatedLineRows(dlm, selectedLineRowId, lineGroupRowId);
                if (lineRowIds.Count == 0)
                {
                    lineRowIds.Add(selectedLineRowId);
                }

                List<EquipmentCandidate> candidates = CollectEquipmentCandidates(dlm, pnpDb, db, lineRowIds)
                    .OrderBy(x => x.Tag, StringComparer.OrdinalIgnoreCase)
                    .ThenBy(x => x.DrawingName, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (candidates.Count < 2)
                {
                    ed.WriteMessage(
                        "\nAt least 2 equipment items need to be connected to the line group for this command.");
                    return;
                }

                EquipmentCandidate fromCandidate;
                EquipmentCandidate toCandidate;

                if (candidates.Count == 2)
                {
                    if (!ResolveByTwoEquipmentConfirmation(ed, lineGroupTag, candidates, out fromCandidate, out toCandidate))
                    {
                        ed.WriteMessage("\nCommand canceled.");
                        return;
                    }
                }
                else
                {
                    ed.WriteMessage(
                        "\nMore than two equipment items were found for line group " + lineGroupTag + ".");

                    fromCandidate = PromptForEndpointCandidate(ed, dlm, pnpDb, candidates, "FROM", null);
                    if (fromCandidate == null)
                    {
                        ed.WriteMessage("\nCommand canceled.");
                        return;
                    }

                    toCandidate = PromptForEndpointCandidate(ed, dlm, pnpDb, candidates, "TO", fromCandidate.RowId);
                    if (toCandidate == null)
                    {
                        ed.WriteMessage("\nCommand canceled.");
                        return;
                    }
                }

                WriteLineGroupFromTo(
                    dlm,
                    lineGroupRowId,
                    actualFromPropertyName,
                    actualToPropertyName,
                    fromCandidate.Tag,
                    toCandidate.Tag);

                ed.WriteMessage(
                    "\nPipe Line Group " +
                    lineGroupTag +
                    ": " +
                    actualFromPropertyName +
                    "='" +
                    fromCandidate.Tag +
                    "', " +
                    actualToPropertyName +
                    "='" +
                    toCandidate.Tag +
                    "'.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\n" + CommandName + " failed: " + ex.Message);
            }
        }

        private static Project GetPidProject()
        {
            PlantProject currentProject = PlantApplication.CurrentProject;
            if (currentProject == null)
            {
                return null;
            }

            try
            {
                return currentProject.ProjectParts["PnId"];
            }
            catch
            {
                // Ignore and try alternate key.
            }

            try
            {
                return currentProject.ProjectParts["PnID"];
            }
            catch
            {
                return null;
            }
        }

        private static ObjectId PromptForPipeLineSegment(Editor ed)
        {
            PromptEntityOptions peo = new PromptEntityOptions("\nSelect a Pipe Line Segment: ");
            peo.SetRejectMessage("\nOnly Pipe Line Segment objects are allowed.");
            peo.AddAllowedClass(typeof(LineSegment), true);

            PromptEntityResult per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
            {
                return ObjectId.Null;
            }

            return per.ObjectId;
        }

        private static int SafeFindRowId(DataLinksManager dlm, ObjectId objectId)
        {
            try
            {
                return dlm.FindAcPpRowId(objectId);
            }
            catch
            {
                return 0;
            }
        }

        private static int GetLineGroupRowId(DataLinksManager dlm, int lineRowId)
        {
            foreach (int groupId in SafeGetRelatedRows(dlm, RelPipeLineGroup, "PipeLine", lineRowId, "PipeLineGroup"))
            {
                if (groupId > 0)
                {
                    return groupId;
                }
            }

            return 0;
        }

        private static HashSet<int> CollectAllRelatedLineRows(
            DataLinksManager dlm,
            int selectedLineRowId,
            int lineGroupRowId)
        {
            HashSet<int> results = new HashSet<int>();
            Queue<int> queue = new Queue<int>();

            AddLineRow(results, queue, selectedLineRowId);

            foreach (int lineRowId in SafeGetRelatedRows(dlm, RelPipeLineGroup, "PipeLineGroup", lineGroupRowId, "PipeLine"))
            {
                AddLineRow(results, queue, lineRowId);
            }

            while (queue.Count > 0)
            {
                int currentLineRowId = queue.Dequeue();

                foreach (int opcRowId in SafeGetRelatedRows(
                    dlm,
                    RelLineOffPageConnector,
                    "Line",
                    currentLineRowId,
                    "OffPageConnector"))
                {
                    foreach (int connectedOpcRowId in GetConnectedOffPageConnectorRowIds(dlm, opcRowId))
                    {
                        foreach (int remoteLineRowId in SafeGetRelatedRows(
                            dlm,
                            RelLineOffPageConnector,
                            "OffPageConnector",
                            connectedOpcRowId,
                            "Line"))
                        {
                            AddLineRow(results, queue, remoteLineRowId);
                        }
                    }
                }
            }

            return results;
        }

        private static void AddLineRow(HashSet<int> results, Queue<int> queue, int lineRowId)
        {
            if (lineRowId <= 0)
            {
                return;
            }

            if (results.Add(lineRowId))
            {
                queue.Enqueue(lineRowId);
            }
        }

        private static IEnumerable<int> GetConnectedOffPageConnectorRowIds(DataLinksManager dlm, int opcRowId)
        {
            HashSet<int> ids = new HashSet<int>();

            foreach (int rowId in SafeGetRelatedRows(dlm, RelConnectors, "Connector1", opcRowId, "Connector2"))
            {
                ids.Add(rowId);
            }

            foreach (int rowId in SafeGetRelatedRows(dlm, RelConnectors, "Connector2", opcRowId, "Connector1"))
            {
                ids.Add(rowId);
            }

            return ids;
        }

        private static List<int> SafeGetRelatedRows(
            DataLinksManager dlm,
            string relationshipName,
            string fromRole,
            int fromRowId,
            string toRole)
        {
            List<int> ids = new List<int>();

            try
            {
                var rowIds = dlm.GetRelatedRowIds(relationshipName, fromRole, fromRowId, toRole);
                foreach (int rowId in rowIds)
                {
                    if (rowId > 0)
                    {
                        ids.Add(rowId);
                    }
                }
            }
            catch
            {
                // Missing relationship or no related rows - return empty list.
            }

            return ids;
        }

        private static List<EquipmentCandidate> CollectEquipmentCandidates(
            DataLinksManager dlm,
            PnPDatabase pnpDb,
            Database db,
            IEnumerable<int> lineRowIds)
        {
            Dictionary<int, EquipmentCandidate> candidatesByRowId =
                new Dictionary<int, EquipmentCandidate>();

            foreach (int lineRowId in lineRowIds)
            {
                HashSet<int> equipmentRowIds = new HashSet<int>();

                AddResolvedEquipmentRows(
                    equipmentRowIds,
                    ResolveEquipmentRowsFromTerminals(
                        dlm,
                        pnpDb,
                        SafeGetRelatedRows(dlm, RelLineStartAsset, "Line", lineRowId, "Asset")));

                AddResolvedEquipmentRows(
                    equipmentRowIds,
                    ResolveEquipmentRowsFromTerminals(
                        dlm,
                        pnpDb,
                        SafeGetRelatedRows(dlm, RelLineEndAsset, "Line", lineRowId, "Asset")));

                AddResolvedEquipmentRows(
                    equipmentRowIds,
                    ResolveEquipmentRowsFromTerminals(
                        dlm,
                        pnpDb,
                        SafeGetRelatedRows(dlm, RelLineNozzle, "Line", lineRowId, "Nozzle")));

                foreach (int equipmentRowId in equipmentRowIds)
                {
                    if (!candidatesByRowId.ContainsKey(equipmentRowId))
                    {
                        EquipmentCandidate candidate = BuildEquipmentCandidate(dlm, pnpDb, db, equipmentRowId);
                        if (candidate != null)
                        {
                            candidatesByRowId.Add(equipmentRowId, candidate);
                        }
                    }
                }
            }

            return candidatesByRowId.Values.ToList();
        }

        private static void AddResolvedEquipmentRows(HashSet<int> target, IEnumerable<int> equipmentRowIds)
        {
            foreach (int equipmentRowId in equipmentRowIds)
            {
                if (equipmentRowId > 0)
                {
                    target.Add(equipmentRowId);
                }
            }
        }

        private static IEnumerable<int> ResolveEquipmentRowsFromTerminals(
            DataLinksManager dlm,
            PnPDatabase pnpDb,
            IEnumerable<int> terminalRowIds)
        {
            HashSet<int> resolvedEquipmentIds = new HashSet<int>();

            foreach (int terminalRowId in terminalRowIds)
            {
                foreach (int equipmentRowId in ResolveEquipmentRowsFromTerminal(dlm, pnpDb, terminalRowId))
                {
                    resolvedEquipmentIds.Add(equipmentRowId);
                }
            }

            return resolvedEquipmentIds;
        }

        private static IEnumerable<int> ResolveEquipmentRowsFromTerminal(
            DataLinksManager dlm,
            PnPDatabase pnpDb,
            int terminalRowId)
        {
            HashSet<int> results = new HashSet<int>();

            if (RowExists(pnpDb, TableEquipment, terminalRowId))
            {
                results.Add(terminalRowId);
                return results;
            }

            if (RowExists(pnpDb, TableNozzles, terminalRowId))
            {
                foreach (int ownerRowId in SafeGetRelatedRows(
                    dlm,
                    RelAssetOwnership,
                    "Owned",
                    terminalRowId,
                    "Owner"))
                {
                    if (RowExists(pnpDb, TableEquipment, ownerRowId))
                    {
                        results.Add(ownerRowId);
                    }
                }
            }

            return results;
        }

        private static bool RowExists(PnPDatabase pnpDb, string tableName, int rowId)
        {
            try
            {
                PnPTable table = pnpDb.Tables[tableName];
                if (table == null)
                {
                    return false;
                }

                PnPRow[] rows = table.Select("\"PnPID\"=" + rowId.ToString());
                return rows != null && rows.Length > 0;
            }
            catch
            {
                return false;
            }
        }

        private static EquipmentCandidate BuildEquipmentCandidate(
            DataLinksManager dlm,
            PnPDatabase pnpDb,
            Database db,
            int equipmentRowId)
        {
            List<KeyValuePair<string, string>> props = dlm.GetAllProperties(equipmentRowId, true);
            string tag = GetPreferredPropertyValue(props, "Tag", "Number", "Name");
            if (string.IsNullOrWhiteSpace(tag))
            {
                tag = "<row " + equipmentRowId.ToString() + ">";
            }

            string drawingName = GetDrawingNameForRowId(pnpDb, equipmentRowId);
            ObjectId objectId = FindObjectIdInCurrentDrawing(dlm, db, equipmentRowId);

            return new EquipmentCandidate(equipmentRowId, tag, drawingName, objectId);
        }

        private static string GetDrawingNameForRowId(PnPDatabase pnpDb, int rowId)
        {
            try
            {
                PnPTable dataLinksTable = pnpDb.Tables[TablePnPDataLinks];
                PnPTable drawingsTable = pnpDb.Tables[TablePnPDrawings];
                if (dataLinksTable == null || drawingsTable == null)
                {
                    return string.Empty;
                }

                PnPRow[] linkRows = dataLinksTable.Select("\"RowId\"=" + rowId.ToString());
                if (linkRows == null)
                {
                    return string.Empty;
                }

                foreach (PnPRow linkRow in linkRows)
                {
                    object dwgIdObj = linkRow["DwgId"];
                    if (dwgIdObj == null)
                    {
                        continue;
                    }

                    int dwgId;
                    try
                    {
                        dwgId = Convert.ToInt32(dwgIdObj);
                    }
                    catch
                    {
                        continue;
                    }

                    PnPRow[] drawingRows = drawingsTable.Select("\"PnPID\"=" + dwgId.ToString());
                    if (drawingRows == null || drawingRows.Length == 0)
                    {
                        continue;
                    }

                    string drawingName = Convert.ToString(drawingRows[0]["Dwg Name"]);
                    if (!string.IsNullOrWhiteSpace(drawingName))
                    {
                        return drawingName;
                    }
                }
            }
            catch
            {
                // Ignore and fall through.
            }

            return string.Empty;
        }

        private static ObjectId FindObjectIdInCurrentDrawing(DataLinksManager dlm, Database db, int rowId)
        {
            try
            {
                var ppObjectIds = dlm.FindAcPpObjectIds(rowId);
                foreach (var ppObjectId in ppObjectIds)
                {
                    try
                    {
                        ObjectId objectId = dlm.MakeAcDbObjectId(ppObjectId);
                        if (!objectId.IsNull && objectId.Database == db)
                        {
                            return objectId;
                        }
                    }
                    catch
                    {
                        // Ignore invalid conversions for drawings that are not loaded.
                    }
                }
            }
            catch
            {
                // Ignore and return null ObjectId.
            }

            return ObjectId.Null;
        }

        private static bool ResolveByTwoEquipmentConfirmation(
            Editor ed,
            string lineGroupTag,
            List<EquipmentCandidate> candidates,
            out EquipmentCandidate fromCandidate,
            out EquipmentCandidate toCandidate)
        {
            fromCandidate = null;
            toCandidate = null;

            EquipmentCandidate first = candidates[0];
            EquipmentCandidate second = candidates[1];
            int noCount = 0;

            while (noCount < 2)
            {
                EquipmentCandidate proposedFrom = noCount == 0 ? first : second;
                EquipmentCandidate proposedTo = noCount == 0 ? second : first;

                ConfirmationChoice choice = PromptSuggestedDirection(
                    ed,
                    lineGroupTag,
                    proposedFrom,
                    proposedTo);

                if (choice == ConfirmationChoice.Yes)
                {
                    fromCandidate = proposedFrom;
                    toCandidate = proposedTo;
                    return true;
                }

                if (choice == ConfirmationChoice.Cancel)
                {
                    return false;
                }

                noCount++;
            }

            return false;
        }

        private static ConfirmationChoice PromptSuggestedDirection(
            Editor ed,
            string lineGroupTag,
            EquipmentCandidate fromCandidate,
            EquipmentCandidate toCandidate)
        {
            string message =
                "\nPipeline group " +
                lineGroupTag +
                " goes from " +
                fromCandidate.GetDisplayText() +
                " to " +
                toCandidate.GetDisplayText() +
                ". Is this correct";

            PromptKeywordOptions pko = new PromptKeywordOptions(message);
            pko.Keywords.Add("Yes");
            pko.Keywords.Add("No");
            pko.Keywords.Add("Cancel");
            pko.Keywords.Default = "Yes";
            pko.AllowNone = true;

            PromptResult pr = ed.GetKeywords(pko);
            if (pr.Status == PromptStatus.Cancel)
            {
                return ConfirmationChoice.Cancel;
            }

            string answer = pr.StringResult ?? string.Empty;
            if (string.IsNullOrWhiteSpace(answer) ||
                answer.Equals("Yes", StringComparison.OrdinalIgnoreCase))
            {
                return ConfirmationChoice.Yes;
            }

            if (answer.Equals("No", StringComparison.OrdinalIgnoreCase))
            {
                return ConfirmationChoice.No;
            }

            return ConfirmationChoice.Cancel;
        }

        private static EquipmentCandidate PromptForEndpointCandidate(
            Editor ed,
            DataLinksManager dlm,
            PnPDatabase pnpDb,
            List<EquipmentCandidate> allCandidates,
            string endpointName,
            int? excludedRowId)
        {
            List<EquipmentCandidate> candidates = new List<EquipmentCandidate>();
            foreach (EquipmentCandidate candidate in allCandidates)
            {
                if (!excludedRowId.HasValue || candidate.RowId != excludedRowId.Value)
                {
                    candidates.Add(candidate);
                }
            }

            if (candidates.Count == 0)
            {
                return null;
            }

            while (true)
            {
                ed.WriteMessage("\nAvailable " + endpointName + " equipment:");
                for (int i = 0; i < candidates.Count; i++)
                {
                    ed.WriteMessage("\n  E" + (i + 1).ToString() + " = " + candidates[i].GetDisplayText());
                }

                bool anyPickable = candidates.Any(x => x.IsOnCurrentDrawing);

                PromptKeywordOptions pko = new PromptKeywordOptions(
                    "\nSelect " + endpointName + " equipment");
                for (int i = 0; i < candidates.Count; i++)
                {
                    pko.Keywords.Add("E" + (i + 1).ToString());
                }

                if (anyPickable)
                {
                    pko.Keywords.Add("Pick");
                }

                pko.Keywords.Add("Cancel");
                pko.AllowNone = false;

                PromptResult pr = ed.GetKeywords(pko);
                if (pr.Status == PromptStatus.Cancel)
                {
                    return null;
                }

                string answer = pr.StringResult ?? string.Empty;
                if (answer.Equals("Cancel", StringComparison.OrdinalIgnoreCase))
                {
                    return null;
                }

                if (answer.Equals("Pick", StringComparison.OrdinalIgnoreCase))
                {
                    bool pickCanceled;
                    EquipmentCandidate picked = PromptPickCandidate(
                        ed,
                        dlm,
                        pnpDb,
                        candidates,
                        endpointName,
                        out pickCanceled);

                    if (pickCanceled)
                    {
                        return null;
                    }

                    if (picked != null)
                    {
                        return picked;
                    }

                    continue;
                }

                if (answer.StartsWith("E", StringComparison.OrdinalIgnoreCase))
                {
                    int index;
                    if (int.TryParse(answer.Substring(1), out index) &&
                        index >= 1 &&
                        index <= candidates.Count)
                    {
                        return candidates[index - 1];
                    }
                }
            }
        }

        private static EquipmentCandidate PromptPickCandidate(
            Editor ed,
            DataLinksManager dlm,
            PnPDatabase pnpDb,
            List<EquipmentCandidate> candidates,
            string endpointName,
            out bool canceled)
        {
            canceled = false;

            List<EquipmentCandidate> pickableCandidates = candidates
                .Where(x => x.IsOnCurrentDrawing && !x.ObjectId.IsNull)
                .ToList();

            if (pickableCandidates.Count == 0)
            {
                ed.WriteMessage("\nNone of the listed equipment is on the active drawing.");
                return null;
            }

            while (true)
            {
                PromptEntityOptions peo = new PromptEntityOptions(
                    "\nClick the " + endpointName + " equipment on the active drawing: ");
                peo.SetRejectMessage("\nSelect one of the listed equipment objects.");
                peo.AddAllowedClass(typeof(Asset), true);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status == PromptStatus.Cancel)
                {
                    canceled = true;
                    return null;
                }

                if (per.Status != PromptStatus.OK)
                {
                    continue;
                }

                int equipmentRowId;
                if (!TryResolveEquipmentRowFromPickedObject(dlm, pnpDb, per.ObjectId, out equipmentRowId))
                {
                    ed.WriteMessage("\nThe selected object is not a listed equipment item.");
                    continue;
                }

                EquipmentCandidate matched = pickableCandidates
                    .FirstOrDefault(x => x.RowId == equipmentRowId);
                if (matched != null)
                {
                    return matched;
                }

                ed.WriteMessage("\nThe selected object is not one of the listed equipment items.");
            }
        }

        private static bool TryResolveEquipmentRowFromPickedObject(
    DataLinksManager dlm,
    PnPDatabase pnpDb,
    ObjectId objectId,
    out int equipmentRowId)
        {
            equipmentRowId = 0;

            int pickedRowId = SafeFindRowId(dlm, objectId);
            if (pickedRowId <= 0)
            {
                return false;
            }

            foreach (int resolvedRowId in ResolveEquipmentRowsFromTerminal(dlm, pnpDb, pickedRowId))
            {
                if (resolvedRowId > 0)
                {
                    equipmentRowId = resolvedRowId;
                    return true;
                }
            }

            return false;
        }

        private static bool TryResolveLineGroupPropertyNames(
            DataLinksManager dlm,
            int lineGroupRowId,
            out string fromPropertyName,
            out string toPropertyName)
        {
            fromPropertyName = string.Empty;
            toPropertyName = string.Empty;

            List<KeyValuePair<string, string>> props = dlm.GetAllProperties(lineGroupRowId, true);
            if (props == null || props.Count == 0)
            {
                return false;
            }

            fromPropertyName = FindActualPropertyName(props, RequestedFromProperty);
            toPropertyName = FindActualPropertyName(props, RequestedToProperty);

            return !string.IsNullOrWhiteSpace(fromPropertyName) &&
                   !string.IsNullOrWhiteSpace(toPropertyName);
        }

        private static string FindActualPropertyName(
            IEnumerable<KeyValuePair<string, string>> props,
            string requestedName)
        {
            foreach (KeyValuePair<string, string> prop in props)
            {
                if (string.Equals(prop.Key, requestedName, StringComparison.OrdinalIgnoreCase))
                {
                    return prop.Key;
                }
            }

            return string.Empty;
        }

        private static string GetLineGroupDisplayTag(DataLinksManager dlm, int lineGroupRowId)
        {
            List<KeyValuePair<string, string>> props = dlm.GetAllProperties(lineGroupRowId, true);
            string value = GetPreferredPropertyValue(props, "Tag", "LineNumber", "Line Number", "Number");
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value;
            }

            return "#" + lineGroupRowId.ToString();
        }

        private static string GetPreferredPropertyValue(
            IEnumerable<KeyValuePair<string, string>> props,
            params string[] preferredNames)
        {
            if (props == null)
            {
                return string.Empty;
            }

            foreach (string preferredName in preferredNames)
            {
                foreach (KeyValuePair<string, string> prop in props)
                {
                    if (string.Equals(prop.Key, preferredName, StringComparison.OrdinalIgnoreCase) &&
                        !string.IsNullOrWhiteSpace(prop.Value))
                    {
                        return prop.Value;
                    }
                }
            }

            return string.Empty;
        }

        private static void WriteLineGroupFromTo(
            DataLinksManager dlm,
            int lineGroupRowId,
            string fromPropertyName,
            string toPropertyName,
            string fromValue,
            string toValue)
        {
            StringCollection names = new StringCollection();
            names.Add(fromPropertyName);
            names.Add(toPropertyName);

            StringCollection values = new StringCollection();
            values.Add(fromValue ?? string.Empty);
            values.Add(toValue ?? string.Empty);

            dlm.SetProperties(lineGroupRowId, names, values);
        }

        private enum ConfirmationChoice
        {
            Yes,
            No,
            Cancel
        }

        private class EquipmentCandidate
        {
            public EquipmentCandidate(int rowId, string tag, string drawingName, ObjectId objectId)
            {
                RowId = rowId;
                Tag = tag ?? string.Empty;
                DrawingName = drawingName ?? string.Empty;
                ObjectId = objectId;
                IsOnCurrentDrawing = !objectId.IsNull;
            }

            public int RowId { get; private set; }

            public string Tag { get; private set; }

            public string DrawingName { get; private set; }

            public bool IsOnCurrentDrawing { get; private set; }

            public ObjectId ObjectId { get; private set; }

            public string GetDisplayText()
            {
                if (!IsOnCurrentDrawing && !string.IsNullOrWhiteSpace(DrawingName))
                {
                    return Tag + " (" + DrawingName + ")";
                }

                return Tag;
            }
        }
    }
}
