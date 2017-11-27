using UnityEngine;

public class ExcelDataInit
{
    public static void Init()
    {
		 Resources.Load<TestConfig_Sheet1>("DataConfig/TestConfig_Sheet1").SetDic();
        Resources.UnloadUnusedAssets();
    }
}