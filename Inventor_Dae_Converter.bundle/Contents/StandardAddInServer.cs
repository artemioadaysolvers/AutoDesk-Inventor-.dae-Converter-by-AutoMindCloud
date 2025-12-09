using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Globalization;
using System.Text;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Reflection;

using Inventor;

// Aliases para evitar conflictos con Inventor.File
using IOPath = System.IO.Path;
using IOFile = System.IO.File;

namespace InventorDaeConverter
{
    [Guid("4aba8fec-e1bf-4786-ba37-6163b7bc0953")]
    [ComVisible(true)]
    public class StandardAddInServer : ApplicationAddInServer
    {
        private Inventor.Application _invApp;
        private ButtonDefinition _exportDaeButton;

        // GUID sin guiones como clientId
        private const string _clientId = "4aba8fece1bf4786ba376163b7bc0953";

        // IDs básicos de material/efecto en el DAE
        private const string _materialId = "mat0";
        private const string _effectId   = "mat0-effect";

        // Flags de debug (puedes cambiar a false/true)
        private static bool _DEBUG_LOG_GENERAL          = true;
        private static bool _DEBUG_LOG_FACETS           = true;
        private static bool _DEBUG_DUMP_OCCURRENCE_MEMS = false;
        private static bool _DEBUG_DUMP_BODY_MEMS       = false;

        // ------------------------------------------------------------
        // OutputDebugString directo → siempre aparece en DebugView
        // ------------------------------------------------------------
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern void OutputDebugString(string lpOutputString);

        private static void Log(string msg)
        {
            try
            {
                OutputDebugString("[DAE] " + msg + "\r\n");
            }
            catch
            {
                // ignorar errores de logging
            }
        }

        // ------------------------------------------------------------
        // Activate: se ejecuta cuando Inventor carga el Add-In
        // ------------------------------------------------------------
        public void Activate(ApplicationAddInSite addInSiteObject, bool firstTime)
        {
            try
            {
                _invApp = addInSiteObject.Application;

                if (_DEBUG_LOG_GENERAL)
                    Log("AddIn Activate. Inventor version=" + _invApp.SoftwareVersion.DisplayVersion);

                CommandManager cmdMgr = _invApp.CommandManager;
                ControlDefinitions controlDefs = cmdMgr.ControlDefinitions;

                // Intentar reutilizar el botón si ya existe
                try
                {
                    object existing = controlDefs["InventorDaeConverter.ExportDae"];
                    _exportDaeButton = existing as ButtonDefinition;
                }
                catch
                {
                    _exportDaeButton = null;
                }

                if (_exportDaeButton == null)
                {
                    _exportDaeButton = controlDefs.AddButtonDefinition(
                        "Exportar DAE",
                        "InventorDaeConverter.ExportDae",
                        CommandTypesEnum.kNonShapeEditCmdType,
                        _clientId,
                        "Exporta el documento activo a COLLADA (.dae).",
                        "Exporta el documento activo a COLLADA (.dae).",
                        Type.Missing,
                        Type.Missing,
                        ButtonDisplayEnum.kDisplayTextInLearningMode);
                }

                _exportDaeButton.OnExecute +=
                    new ButtonDefinitionSink_OnExecuteEventHandler(OnExportDaeButtonExecute);

                // Igual que el HelloWorld: siempre nos aseguramos de que hay UI
                AddToUserInterface();
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al activar el AddIn DAE:\n" + ex.Message,
                    "Inventor DAE Converter");
            }
        }

        // ------------------------------------------------------------
        // Añadir el botón a las cintas (ZeroDoc, Part, Assembly)
        // ------------------------------------------------------------
        private void AddToUserInterface()
        {
            try
            {
                UserInterfaceManager uiMgr = _invApp.UserInterfaceManager;

                // ZeroDoc
                try
                {
                    Ribbon rZero = uiMgr.Ribbons["ZeroDoc"];
                    RibbonTab tabZero = rZero.RibbonTabs["id_TabTools"];

                    RibbonPanel panelZero = null;
                    try
                    {
                        panelZero = tabZero.RibbonPanels["InventorDaeConverter.Panel.ZeroDoc"];
                    }
                    catch
                    {
                        panelZero = null;
                    }

                    if (panelZero == null)
                    {
                        panelZero = tabZero.RibbonPanels.Add(
                            "DAE Export",
                            "InventorDaeConverter.Panel.ZeroDoc",
                            _clientId,
                            "",
                            false);
                    }

                    panelZero.CommandControls.AddButton(_exportDaeButton);
                }
                catch
                {
                    // Ignoramos errores de ZeroDoc
                }

                // Part
                try
                {
                    Ribbon rPart = uiMgr.Ribbons["Part"];
                    RibbonTab tabPart = rPart.RibbonTabs["id_TabTools"];

                    RibbonPanel panelPart = null;
                    try
                    {
                        panelPart = tabPart.RibbonPanels["InventorDaeConverter.Panel.Part"];
                    }
                    catch
                    {
                        panelPart = null;
                    }

                    if (panelPart == null)
                    {
                        panelPart = tabPart.RibbonPanels.Add(
                            "DAE Export",
                            "InventorDaeConverter.Panel.Part",
                            _clientId,
                            "",
                            false);
                    }

                    panelPart.CommandControls.AddButton(_exportDaeButton);
                }
                catch
                {
                    // Ignoramos errores de Part
                }

                // Assembly
                try
                {
                    Ribbon rAsm = uiMgr.Ribbons["Assembly"];
                    RibbonTab tabAsm = rAsm.RibbonTabs["id_TabTools"];

                    RibbonPanel panelAsm = null;
                    try
                    {
                        panelAsm = tabAsm.RibbonPanels["InventorDaeConverter.Panel.Assembly"];
                    }
                    catch
                    {
                        panelAsm = null;
                    }

                    if (panelAsm == null)
                    {
                        panelAsm = tabAsm.RibbonPanels.Add(
                            "DAE Export",
                            "InventorDaeConverter.Panel.Assembly",
                            _clientId,
                            "",
                            false);
                    }

                    panelAsm.CommandControls.AddButton(_exportDaeButton);
                }
                catch
                {
                    // Ignoramos errores de Assembly
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al añadir el botón DAE a la interfaz:\n" + ex.Message,
                    "Inventor DAE Converter");
            }
        }

        // ------------------------------------------------------------
        // Evento del botón: exportar a .dae
        // ------------------------------------------------------------
        private void OnExportDaeButtonExecute(NameValueMap context)
        {
            try
            {
                Document doc = _invApp.ActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("No hay documento activo.", "Inventor DAE Converter");
                    return;
                }

                if (_DEBUG_LOG_GENERAL)
                {
                    Log("OnExportDaeButtonExecute - Documento activo:");
                    Log("  Name=" + doc.DisplayName);
                    Log("  FullFileName=" + doc.FullFileName);
                    Log("  DocumentType=" + doc.DocumentType.ToString());
                }

                if (doc.DocumentType != DocumentTypeEnum.kAssemblyDocumentObject &&
                    doc.DocumentType != DocumentTypeEnum.kPartDocumentObject)
                {
                    MessageBox.Show("Sólo se soportan documentos de pieza (.ipt) y ensamblaje (.iam).",
                        "Inventor DAE Converter");
                    return;
                }

                string defaultName = IOPath.ChangeExtension(doc.DisplayName, ".dae");
                string folder = IOPath.GetDirectoryName(doc.FullFileName);
                if (string.IsNullOrEmpty(folder))
                {
                    folder = System.Environment.GetFolderPath(
                        System.Environment.SpecialFolder.MyDocuments);
                }

                SaveFileDialog sfd = new SaveFileDialog();
                sfd.Title = "Guardar como DAE";
                sfd.InitialDirectory = folder;
                sfd.FileName = defaultName;
                sfd.Filter = "COLLADA DAE (*.dae)|*.dae";

                DialogResult dr = sfd.ShowDialog();
                if (dr != DialogResult.OK || string.IsNullOrEmpty(sfd.FileName))
                {
                    return; // usuario canceló
                }

                string outPath = sfd.FileName;

                // Construimos el DAE
                StringBuilder dae = new StringBuilder();

                dae.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                dae.AppendLine("<COLLADA xmlns=\"http://www.collada.org/2005/11/COLLADASchema\" version=\"1.4.1\">");
                dae.AppendLine("  <asset>");
                dae.AppendLine("    <contributor><authoring_tool>Inventor DAE Converter</authoring_tool></contributor>");
                dae.AppendLine("    <unit name=\"meter\" meter=\"1\"/>");
                dae.AppendLine("    <up_axis>Z_UP</up_axis>");
                dae.AppendLine("  </asset>");

                // Geometría
                dae.AppendLine("  <library_geometries>");

                int meshIndex = 0;

                if (doc.DocumentType == DocumentTypeEnum.kPartDocumentObject)
                {
                    PartDocument partDoc = (PartDocument)doc;
                    PartComponentDefinition defPart = partDoc.ComponentDefinition;

                    SurfaceBodies bodies = defPart.SurfaceBodies;
                    int bodyCount = bodies.Count;

                    if (_DEBUG_LOG_GENERAL)
                    {
                        Log("PartDocument: SurfaceBodies.Count=" + bodyCount);
                    }

                    foreach (SurfaceBody body in bodies)
                    {
                        try
                        {
                            if (!body.Visible)
                            {
                                if (_DEBUG_LOG_GENERAL)
                                    Log("  Body invisible -> skip");
                                continue;
                            }

                            string geomName = "PartBody_" + meshIndex.ToString(CultureInfo.InvariantCulture);

                            if (_DEBUG_DUMP_BODY_MEMS)
                            {
                                DumpComMembers(body, "SurfaceBody (Part)", null);
                                DumpComMembers(body, "SurfaceBody (Facet)", "Facet");
                                DumpComMembers(body, "SurfaceBody (Tessell)", "Tessell");
                            }

                            ExportSurfaceBodyToDae(body, null, geomName, ref dae, ref meshIndex);
                        }
                        catch (Exception exB)
                        {
                            Log("Error exportando SurfaceBody (part): " + exB.Message);
                        }
                    }
                }
                else // Assembly
                {
                    AssemblyDocument asmDoc = (AssemblyDocument)doc;
                    AssemblyComponentDefinition defAsm = asmDoc.ComponentDefinition;

                    if (_DEBUG_LOG_GENERAL)
                    {
                        Log("AssemblyDocument: Occurrences.Count=" + defAsm.Occurrences.Count);
                    }

                    ExportOccurrencesRecursive(defAsm.Occurrences, ref dae, ref meshIndex);
                }

                dae.AppendLine("  </library_geometries>");

                // Efecto/material mínimos (para que los visores no ignoren la geometría)
                dae.AppendLine("  <library_effects>");
                dae.AppendLine("    <effect id=\"" + _effectId + "\">");
                dae.AppendLine("      <profile_COMMON>");
                dae.AppendLine("        <technique sid=\"common\">");
                dae.AppendLine("          <phong>");
                dae.AppendLine("            <diffuse>");
                dae.AppendLine("              <color>0.8 0.8 0.8 1</color>");
                dae.AppendLine("            </diffuse>");
                dae.AppendLine("          </phong>");
                dae.AppendLine("        </technique>");
                dae.AppendLine("      </profile_COMMON>");
                dae.AppendLine("    </effect>");
                dae.AppendLine("  </library_effects>");

                dae.AppendLine("  <library_materials>");
                dae.AppendLine("    <material id=\"" + _materialId + "\" name=\"" + _materialId + "\">");
                dae.AppendLine("      <instance_effect url=\"#" + _effectId + "\"/>");
                dae.AppendLine("    </material>");
                dae.AppendLine("  </library_materials>");

                // Escena: un nodo por geometría, con bind_material
                dae.AppendLine("  <library_visual_scenes>");
                dae.AppendLine("    <visual_scene id=\"Scene\" name=\"Scene\">");
                for (int i = 0; i < meshIndex; ++i)
                {
                    string geomId = "geom_" + i.ToString(CultureInfo.InvariantCulture);
                    string nodeId = "node_" + i.ToString(CultureInfo.InvariantCulture);

                    dae.AppendLine("      <node id=\"" + nodeId + "\" name=\"" + nodeId + "\">");
                    dae.AppendLine("        <instance_geometry url=\"#" + geomId + "\">");
                    dae.AppendLine("          <bind_material>");
                    dae.AppendLine("            <technique_common>");
                    dae.AppendLine("              <instance_material symbol=\"" + _materialId + "\" target=\"#" + _materialId + "\"/>");
                    dae.AppendLine("            </technique_common>");
                    dae.AppendLine("          </bind_material>");
                    dae.AppendLine("        </instance_geometry>");
                    dae.AppendLine("      </node>");
                }
                dae.AppendLine("    </visual_scene>");
                dae.AppendLine("  </library_visual_scenes>");

                dae.AppendLine("  <scene>");
                dae.AppendLine("    <instance_visual_scene url=\"#Scene\"/>");
                dae.AppendLine("  </scene>");

                dae.AppendLine("</COLLADA>");

                IOFile.WriteAllText(outPath, dae.ToString(), Encoding.UTF8);

                string extraNote = meshIndex == 0
                    ? "\nOJO: se exportaron 0 meshes. El visor dirá que el modelo no contiene mallas."
                    : "";

                if (_DEBUG_LOG_GENERAL)
                {
                    Log("Export finalizado. meshIndex=" + meshIndex);
                    Log("Archivo: " + outPath);
                }

                string message =
                    "Exportado a DAE (" +
                    meshIndex.ToString(CultureInfo.InvariantCulture) +
                    " mesh(es))." +
                    extraNote +
                    "\n\n" +
                    outPath;

                MessageBox.Show(
                    message,
                    "Inventor DAE Converter");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al exportar a DAE:\n" + ex.Message, "Inventor DAE Converter");
            }
        }

        // ------------------------------------------------------------
        // Recorrido recursivo del ensamblaje
        // ------------------------------------------------------------
        private void ExportOccurrencesRecursive(IEnumerable occs, ref StringBuilder dae, ref int meshIndex)
        {
            foreach (object o in occs)
            {
                ComponentOccurrence occ = o as ComponentOccurrence;
                if (occ == null)
                    continue;

                try
                {
                    if (occ.Suppressed)
                    {
                        if (_DEBUG_LOG_GENERAL)
                            Log("Occurrence '" + occ.Name + "' suprimida -> skip");
                        continue;
                    }
                    if (!occ.Visible)
                    {
                        if (_DEBUG_LOG_GENERAL)
                            Log("Occurrence '" + occ.Name + "' invisible -> skip");
                        continue;
                    }

                    if (_DEBUG_LOG_GENERAL)
                        Log("Occurrence: " + occ.Name +
                            "  SubOccs=" + occ.SubOccurrences.Count +
                            "  SurfaceBodies=" + occ.SurfaceBodies.Count);

                    if (_DEBUG_DUMP_OCCURRENCE_MEMS)
                    {
                        DumpComMembers(occ, "ComponentOccurrence " + occ.Name, null);
                    }

                    // Subensamblaje → recursivo
                    if (occ.SubOccurrences.Count > 0)
                    {
                        ExportOccurrencesRecursive(occ.SubOccurrences, ref dae, ref meshIndex);
                    }
                    else
                    {
                        // Componente hoja: exportar sus SurfaceBodies
                        SurfaceBodies bodies = occ.SurfaceBodies;
                        foreach (SurfaceBody body in bodies)
                        {
                            if (!body.Visible)
                            {
                                if (_DEBUG_LOG_GENERAL)
                                    Log("  Body invisible en '" + occ.Name + "' -> skip");
                                continue;
                            }

                            string safeOccName = MakeSafeName(occ.Name);
                            string geomName = safeOccName + "_" + meshIndex.ToString(CultureInfo.InvariantCulture);

                            if (_DEBUG_DUMP_BODY_MEMS)
                            {
                                DumpComMembers(body, "SurfaceBody (" + occ.Name + ")", null);
                                DumpComMembers(body, "SurfaceBody (" + occ.Name + ") Facet", "Facet");
                                DumpComMembers(body, "SurfaceBody (" + occ.Name + ") Tessell", "Tessell");
                            }

                            ExportSurfaceBodyToDae(body, occ, geomName, ref dae, ref meshIndex);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("Error ExportOccurrencesRecursive en '" + occ.Name + "': " + ex.Message);
                }
            }
        }

        // ------------------------------------------------------------
        // Sanitizar nombres para IDs en DAE
        // ------------------------------------------------------------
        private string MakeSafeName(string name)
        {
            if (string.IsNullOrEmpty(name))
                return "body";

            StringBuilder sb = new StringBuilder(name.Length);
            foreach (char c in name)
            {
                if (char.IsLetterOrDigit(c) || c == '_')
                    sb.Append(c);
                else
                    sb.Append('_');
            }
            return sb.ToString();
        }

        // ------------------------------------------------------------
        // Exportar un SurfaceBody a geometría DAE usando CalculateFacets
        // (arrays inicializados para evitar DISP_E_TYPEMISMATCH)
        // ------------------------------------------------------------
        private void ExportSurfaceBodyToDae(
            SurfaceBody body,
            ComponentOccurrence occ,
            string geomName,
            ref StringBuilder dae,
            ref int meshIndex)
        {
            try
            {
                // Tolerancia en centímetros (API de Inventor usa cm)
                double tol = 0.1;

                int vertexCount  = 0;
                int facetCount   = 0;

                // *** IMPORTANTE: inicializar arrays vacíos ***
                double[] vertexCoords   = new double[] { };
                double[] normalVectors  = new double[] { };
                int[]    vertexIndices  = new int[] { };

                body.CalculateFacets(
                    tol,
                    out vertexCount,
                    out facetCount,
                    out vertexCoords,
                    out normalVectors,
                    out vertexIndices);

                if (_DEBUG_LOG_FACETS)
                {
                    string occName = (occ != null) ? occ.Name : "<PartRoot>";
                    Log(
                        "CalculateFacets en geom '" + geomName +
                        "' (occ=" + occName +
                        "): vertexCount=" + vertexCount +
                        ", facetCount=" + facetCount +
                        ", vertsLen=" + (vertexCoords == null ? 0 : vertexCoords.Length) +
                        ", idxLen=" + (vertexIndices == null ? 0 : vertexIndices.Length));
                }

                if (vertexCount <= 0 || facetCount <= 0 ||
                    vertexCoords == null || vertexIndices == null)
                {
                    if (_DEBUG_LOG_FACETS)
                    {
                        Log("  Sin facetas -> no se genera geometría para '" + geomName + "'");
                    }
                    return;
                }

                // Convertir índices a base 0 (CalculateFacets devuelve 1-based)
                int triCountTimes3 = facetCount * 3;
                if (vertexIndices.Length < triCountTimes3)
                {
                    triCountTimes3 = vertexIndices.Length;
                }

                int[] tris = new int[triCountTimes3];
                for (int i = 0; i < triCountTimes3; ++i)
                {
                    int idx = vertexIndices[i] - 1;
                    if (idx < 0) idx = 0;
                    tris[i] = idx;
                }

                string geomId = "geom_" + meshIndex.ToString(CultureInfo.InvariantCulture);

                WriteRawMeshToDae(vertexCoords, tris, geomId, geomName, ref dae);

                meshIndex++;
            }
            catch (Exception ex)
            {
                Log("Error ExportSurfaceBodyToDae: " + ex.Message);
            }
        }

        // ------------------------------------------------------------
        // Escribir un mesh crudo (verts + tris) en COLLADA
        // ------------------------------------------------------------
        private void WriteRawMeshToDae(
            double[] verts,
            int[] tris,
            string geomId,
            string geomName,
            ref StringBuilder dae)
        {
            if (verts == null || tris == null)
                return;
            if (verts.Length < 3 || tris.Length < 3)
                return;

            int vertexCount = verts.Length / 3;
            int triangleCount = tris.Length / 3;

            dae.AppendLine("    <geometry id=\"" + geomId + "\" name=\"" + geomName + "\">");
            dae.AppendLine("      <mesh>");

            // Positions
            dae.AppendLine("        <source id=\"" + geomId + "-positions\">");
            dae.AppendLine("          <float_array id=\"" + geomId + "-positions-array\" count=\"" +
                           (vertexCount * 3).ToString(CultureInfo.InvariantCulture) + "\">");

            // verts están en cm → convertir a metros
            for (int i = 0; i < verts.Length; i++)
            {
                double vCm = verts[i];
                double vMeters = vCm * 0.01;
                dae.AppendFormat(CultureInfo.InvariantCulture, "{0} ", vMeters);
            }

            dae.AppendLine("</float_array>");
            dae.AppendLine("          <technique_common>");
            dae.AppendLine("            <accessor source=\"#" + geomId + "-positions-array\" count=\"" +
                           vertexCount.ToString(CultureInfo.InvariantCulture) + "\" stride=\"3\">");
            dae.AppendLine("              <param name=\"X\" type=\"float\"/>");
            dae.AppendLine("              <param name=\"Y\" type=\"float\"/>");
            dae.AppendLine("              <param name=\"Z\" type=\"float\"/>");
            dae.AppendLine("            </accessor>");
            dae.AppendLine("          </technique_common>");
            dae.AppendLine("        </source>");

            // Vertices
            dae.AppendLine("        <vertices id=\"" + geomId + "-vertices\">");
            dae.AppendLine("          <input semantic=\"POSITION\" source=\"#" + geomId + "-positions\"/>");
            dae.AppendLine("        </vertices>");

            // Triangles: con material="mat0"
            dae.AppendLine("        <triangles material=\"" + _materialId + "\" count=\"" +
                           triangleCount.ToString(CultureInfo.InvariantCulture) + "\">");
            dae.AppendLine("          <input semantic=\"VERTEX\" source=\"#" + geomId + "-vertices\" offset=\"0\"/>");
            dae.AppendLine("          <p>");

            for (int i = 0; i < tris.Length; i++)
            {
                dae.AppendFormat(CultureInfo.InvariantCulture, "{0} ", tris[i]);
            }

            dae.AppendLine("</p>");
            dae.AppendLine("        </triangles>");

            dae.AppendLine("      </mesh>");
            dae.AppendLine("    </geometry>");
        }

        // ------------------------------------------------------------
        // Inspector por reflection: lista propiedades y métodos
        // ------------------------------------------------------------
        private void DumpComMembers(object obj, string label, string nameFilter)
        {
            if (obj == null)
            {
                Log("[" + label + "] <null>");
                return;
            }

            Type t = obj.GetType();
            Log("======================================");
            Log("[" + label + "] Type = " + t.FullName);
            Log("---- PROPERTIES ----");

            PropertyInfo[] props = t.GetProperties(BindingFlags.Instance | BindingFlags.Public);
            foreach (PropertyInfo p in props)
            {
                if (!string.IsNullOrEmpty(nameFilter) &&
                    p.Name.IndexOf(nameFilter, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                Log("P: " + p.PropertyType.Name + " " + p.Name);
            }

            Log("---- METHODS ----");

            MethodInfo[] methods = t.GetMethods(BindingFlags.Instance | BindingFlags.Public);
            foreach (MethodInfo m in methods)
            {
                if (m.IsSpecialName)   // getters/setters de propiedades
                    continue;

                if (!string.IsNullOrEmpty(nameFilter) &&
                    m.Name.IndexOf(nameFilter, StringComparison.OrdinalIgnoreCase) < 0)
                    continue;

                ParameterInfo[] pars = m.GetParameters();
                string[] parts = new string[pars.Length];
                for (int i = 0; i < pars.Length; i++)
                {
                    parts[i] = pars[i].ParameterType.Name + " " + pars[i].Name;
                }
                string parsSig = string.Join(", ", parts);

                Log("M: " + m.ReturnType.Name + " " + m.Name + "(" + parsSig + ")");
            }
        }

        // ------------------------------------------------------------
        // Métodos requeridos por ApplicationAddInServer
        // ------------------------------------------------------------
        public void Deactivate()
        {
            try
            {
                if (_exportDaeButton != null)
                {
                    _exportDaeButton.OnExecute -=
                        new ButtonDefinitionSink_OnExecuteEventHandler(OnExportDaeButtonExecute);

                    try
                    {
                        Marshal.FinalReleaseComObject(_exportDaeButton);
                    }
                    catch
                    {
                    }

                    _exportDaeButton = null;
                }

                if (_invApp != null)
                {
                    try
                    {
                        Marshal.FinalReleaseComObject(_invApp);
                    }
                    catch
                    {
                    }
                    _invApp = null;
                }
            }
            catch
            {
            }
            finally
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        public void ExecuteCommand(int CommandID)
        {
            // Obsoleto, no se usa
        }

        public object Automation
        {
            get { return null; }
        }
    }
}
