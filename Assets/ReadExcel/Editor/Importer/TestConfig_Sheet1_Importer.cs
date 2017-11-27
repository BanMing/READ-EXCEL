using UnityEngine;
using System.Collections;
using System.IO;
using UnityEditor;
using System.Xml.Serialization;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.UserModel;
using System.Collections.Generic;
using System.Text;

public class TestConfig_Sheet1_Importer : AssetPostprocessor
{
    private static readonly string filePath = "Assets/ReadExcel/Config/Excel/TestConfig.xlsx";
    private static readonly string sheetName = "Sheet1" ;
    

    static void OnPostprocessAllAssets(string[] importedAssets, string[] deletedAssets, string[] movedAssets, string[] movedFromAssetPaths)
    {
        foreach (string asset in importedAssets)
        {
            if (!filePath.Equals(asset))
                continue;

            using (FileStream stream = File.Open (filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
				IWorkbook book = null;
				if (Path.GetExtension (filePath) == ".xls") {
					book = new HSSFWorkbook(stream);
				} else {
					book = new XSSFWorkbook(stream);
				}


					Directory.CreateDirectory("Assets/ReadExcel/Config/Resources/DataConfig/ ");
                    var exportPath = "Assets/ReadExcel/Config/Resources/DataConfig/TestConfig_Sheet1" + ".asset";
                    
                    // check scriptable object
                    var data = (TestConfig_Sheet1)AssetDatabase.LoadAssetAtPath(exportPath, typeof(TestConfig_Sheet1));
                    if (data == null)
                    {
                        data = ScriptableObject.CreateInstance<TestConfig_Sheet1>();
                        AssetDatabase.CreateAsset((ScriptableObject)data, exportPath);
                        data.hideFlags = HideFlags.NotEditable;
                    }
                    data.Params.Clear();
					data.hideFlags = HideFlags.NotEditable;

					//Creat DataInit Script
					ReadExcel.CreatDataInitCs();

					// check sheet
                    var sheet = book.GetSheet(sheetName);
                    if (sheet == null)
                    {
                        Debug.LogError("[QuestData] sheet not found:" + sheetName);
                        continue;
                    }

                	// add infomation
                    for (int i=2; i<= sheet.LastRowNum; i++)
                    {
                        IRow row = sheet.GetRow(i);
                        ICell cell = null;
                        
                        var p = new TestConfig_Sheet1.Param();
			
                        cell = row.GetCell(0); p.id = (int)(cell == null ? 0 : cell.NumericCellValue);
                        cell = row.GetCell(1); p.name = (cell == null ? "" : cell.StringCellValue);
                        cell = row.GetCell(2); p.age = (int)(cell == null ? 0 : cell.NumericCellValue);
                        p.ts = new List<int>();                        cell = row.GetCell(3);var tstemp = (cell == null ? "" : cell.StringCellValue);string [] tstemps=tstemp.Split(',');for(int index =0,iMax=tstemps.Length;index<iMax;index++){int val=0; val=int.Parse(tstemps[index]);p.ts.Add(val);}
                        p.m = new List<string>();                        cell = row.GetCell(4);var mtemp = (cell == null ? "" : cell.StringCellValue);string [] mtemps=mtemp.Split(',');for(int index =0,iMax=mtemps.Length;index<iMax;index++){string val=mtemps[index];p.m.Add(val);}

                        data.Params.Add(p);
                    }
                    
                    // save scriptable object
                    ScriptableObject obj = AssetDatabase.LoadAssetAtPath(exportPath, typeof(ScriptableObject)) as ScriptableObject;
                    EditorUtility.SetDirty(obj);
                }
            

        }
    }
}
