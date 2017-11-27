#READ-EXCEL
>**写在前面**
这个工具在年初的时候就有一个想法雏形，但是为什么到今天才完成呢，其实就是**懒**，每天下班回家根本都不想动了，就想葛优瘫~这两天在学校度假就捣鼓捣鼓了一下这个东西。最开始想做这个是因为公司的那个读表的系统有个硬伤就是你每改一个表里的一个数据就需要重新全部手动（专门有一个小软件来生成数据代码）生成全部数据脚本，但是这样做应该是像让客户端和服务器用同一套数据脚本，然后就几经查阅就产生了现在这一套工具。

**是什么？**
![工具视图](http://upload-images.jianshu.io/upload_images/3438059-9b338eef226a6b8f.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
这一个工具是读取Excel表中的数据，并把它存在一个ScriptObject数据类。
对于ScriptObject可以看看我的[这一篇文章。](http://www.jianshu.com/p/d2bf2b9436b4)
****

**主要功能**

1.就是上面说的，转化为ScriptObject数据类，这里分为两种：

a.分离一个Excel表中的所有sheet为单独的一个数据体。

b.把一个Excel表中所有的sheet合成在一个数据体中。

**注意：第二种exce中必须每一个sheet的数据项是一样并且位置也是相同的**

2.可以选择每一个sheet或者sheet中的数据单项是否被生成到数据体中。

3.系统会自动识别每一个数据单项的数据类型与数据参数名，当然也可以用户更改。

4.在创建数据体完成后，再改表中的数据，数据体自动更新。当然如果是增加删除单项数据，就需要重新生成数据体。
***

![规则示例](http://upload-images.jianshu.io/upload_images/3438059-197cfcb36242faef.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)
**Excel表格填写规则**

1.第一行为变量名称，数组参数名称结尾为[]。建议不要再次手动修改，不然后面不好区分。

2.第一列必须为id名称，并且必须类型为整形。id名称不要修改。

3.第二行为注释。

4.下面即为数据了。

**注意：这个格式必须按照这样来，不然会出错。**



****
**使用方法**


1.在Project视窗下选中需要生成数据类的Excel表，右键菜单，选择“Read Excel”。

![第一步](http://upload-images.jianshu.io/upload_images/3438059-28879c82e3d42a56.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

2.设置数据类的名字，是否分离，包括哪一些sheet与数据项，检测数据项类型名字。

**注意：浮点型的数需要用户自己设置**
![](http://upload-images.jianshu.io/upload_images/3438059-0f7f304efb89f0da.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

3.所有设置完成即可点击“Create”。

4.生成数据实体类与ScriptObject导入类，会有提示框提示，再次选中Excel表右键点击Reimprot按钮，导入ScriptObject数据。到这一步，编辑器下操作步骤就完成了。


![Paste_Image.png](http://upload-images.jianshu.io/upload_images/3438059-7d0930fe6de43e77.png?imageMogr2/auto-orient/strip%7CimageView2/2/w/1240)

5.然后在游戏载入时调用一下下面这一句话。

    ExcelDataInit.Init();


6.实际运用时数据分为两类：


a.分离单数一个sheet的数据体访问。

    testObj_Sheet1.GetData(1).name


testObj_Sheet1是你的数据类，然后使用GetData方法传入id即可返回数据体实例。


b.合并sheet的数据体访问方式。
    
    testObj.GetSheet((int)testObj.SheetName.allone).GetData(1).ddd


testObj是你的数据类，然后使用GetSheet方法获得sheet，这里自动帮用户生成了sheet的枚举，用户可以直接使用枚举来获得sheet，然后的方法和上一种一样。
*****

**原理**


unity使用.net的api是2.0的，所以没有可以直接读取Excel哪部分源码。就选择使用了NPOI这样一个开源框架。这里有几个学习NPOI的传送门，[官网](http://npoi.codeplex.com/)、[官网繁体文档](https://dotblogs.com.tw/killysss/archive/2010/01/27/13344.aspx)、[官网英文文档](http://www.leniel.net/2014/01/npoi-2.0-major-features-enhancements-series-of-posts-scheduled.html)、[CSDN博客推荐看这个](http://blog.csdn.net/pan_junbiao/article/details/39717443)，具体的就不详细讲解了。

读取到数据后再使用try测试数据类型，自动填充数据类型。这里就是为什么浮点型需要用户自己设置的原因。

使用的ScriptObject这个Unity给出可以用来存数据的资源。这个其实也可以更改成Json文件，这样服务器跟客户端都可以用了。不过ScriptObject可以很直观的看到数据的储存。

这里用到一个资源变化的回调OnPostprocessAllAssets，这个是需要继承于AssetPostprocessor这类的，[详细可以看这里](https://docs.unity3d.com/ScriptReference/AssetPostprocessor.OnPostprocessAllAssets.html)。

这里因为字典不能序列化所以需要在游戏加载时，初始化数据一次。

原理就说这么多吧，再具体的大家可以看看源码。

>**写最后**
感谢各位看官看到最后，这个工具算是写完了，其中的BUG还未有全部测出，优化也没有太多的优化，大家见谅见谅。[博客地址](http://www.jianshu.com/p/f62d2dab8a0a)
