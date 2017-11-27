using UnityEngine;
using System.Collections;
using System.Collections.Generic;

public class TestConfig_Sheet1 : ScriptableObject
{	
	public List<Param> Params=new List<Param>();
    private static Dictionary<int, Param> dicParams = new Dictionary<int, Param>();

	[System.SerializableAttribute]
	public class Param
	{
		
        /// <summary>
        /// 编号
        /// </summary>
        public  int id;

        /// <summary>
        /// 名字
        /// </summary>
        public  string name;

        /// <summary>
        /// 年龄
        /// </summary>
        public  int age;

        /// <summary>
        /// int数组测试
        /// </summary>
        public  List<int> ts;

        /// <summary>
        /// string数组测试
        /// </summary>
        public  List<string> m;

	}

	 /// <summary>
    /// get one data
    /// </summary>
    /// <param name="id">id</param>
    /// <returns></returns>
    public static Param GetData(int id )
    {
        if (dicParams.ContainsKey(id))
        {
            return dicParams[id];
        }
        UnityEngine.Debug.LogError("No This id!");
        return null;
    }


    public void SetDic()
    {
        dicParams.Clear();
        for (int i = 0,iMax = Params.Count; i < iMax ; i++)
        {
            dicParams.Add(Params[i].id, Params[i]);
        }
    }
}