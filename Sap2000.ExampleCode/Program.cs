using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAP2000v20;

namespace Sap2000.ExampleCode
{
    class Program
    {
        static void Main(string[] args)
        {
            //是否启动一个新的sap2000实例
            bool AttachToInstance = false;
            //是否从指定的sap2000启动一个实例，如果不指定，则选择最新安装的版本启动，主要是计算机安装多个sap2000版本的时候 用来选择启动版本
            bool SpecifyPath = false;
            //如果specityPath = true 则通过ProgramPath 指定启动路径
            string ProgramPath = "C:\\Program Files (x86)\\Computers and Structures\\SAP2000 19\\SAP2000.exe";

            //模型的保存文件夹 运行完成后 在这个目录下将会看到结果文件
            string ModelDirectory = "C:\\CSiAPIexample";
            try
            {
                Directory.CreateDirectory(ModelDirectory);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Could not create directory: " + ModelDirectory);
            }
            //模型名称
            string ModelName = "API_1-001.sdb";
            //模型保存路径
            string ModelPath = ModelDirectory + Path.DirectorySeparatorChar + ModelName;

            //定义sap2000实例对象为cOAPI类型 先赋一个空值 后面会赋上实际的对象
            cOAPI mySapObject = null;
            //定义一个变量ret标记函数是佛运行成功 ret=0 为成功 否则为不成功
            int ret = 0;
            //如果AttachToInstance = true 则获取系统运行时中的sap2000实例
            if (AttachToInstance)
            {
                try
                {
                    //获取系统中处于active状态的sap2000实例
                    mySapObject = (cOAPI)System.Runtime.InteropServices.Marshal.GetActiveObject("CSI.SAP2000.API.SapObject");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("No running instance of the program found or failed to attach.");
                    return;
                }
            }
            else
            {//启动新的SAP000实例

                //创建APIhelper对象 该对象可以创建sap2000实例
                cHelper myHelper;
                try
                {
                    myHelper = new Helper();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Cannot create an instance of the Helper object");
                    return;
                }
                //如果specitypath=true 则从programpath创建一个新sap2000实例
                if (SpecifyPath)
                {
                    try
                    {
                        mySapObject = myHelper.CreateObject(ProgramPath);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Cannot start a new instance of the program from " + ProgramPath);
                        return;
                    }
                }
                else
                {
                    //从最新安装的sap2000版本创建实例
                    try
                    {
                        mySapObject = myHelper.CreateObjectProgID("CSI.SAP2000.API.SapObject");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Cannot start a new instance of the program.");
                        return;
                    }
                }
                //启动sap2000实例
                ret = mySapObject.ApplicationStart();
            }

            //创建实例的模型
            cSapModel mySapModel = mySapObject.SapModel;
            //初始化模型 eUnits 变量是什么没搞清楚 应该是模型的类型 ？？？？
            ret = mySapModel.InitializeNewModel((eUnits.kip_in_F));

            //创建一个新的空模型文件 用户保存模型
            ret = mySapModel.File.NewBlank();

            //定义材料属性
            ret = mySapModel.PropMaterial.SetMaterial("CONC", eMatType.Concrete, -1, "", "");
            //各向同性的属性赋给材料
            ret = mySapModel.PropMaterial.SetMPIsotropic("CONC", 3600, 0.2, 0.0000055, 0);
            //定义一个矩形截面并将上面定义的材料赋给截面
            ret = mySapModel.PropFrame.SetRectangle("R1", "CONC", 12, 12, -1, "", "");

            //define frame section property modifiers
            //定义截面的修正值 这个我没搞清楚是什么 一个截面8个值
            double[] ModValue = new double[8];
            int i;
            for (i = 0; i <= 7; i++)
            {
                ModValue[i] = 1;
            }
            ModValue[0] = 1000;
            ModValue[1] = 0;
            ModValue[2] = 0;
            ret = mySapModel.PropFrame.SetModifiers("R1", ref ModValue);

            //设置成k-ft单位制
            ret = mySapModel.SetPresentUnits(eUnits.kip_ft_F);
            //定义线框模型 3条线 Global为坐标系 temp_string1 为系统生成的名称
            string[] FrameName = new string[3];
            string temp_string1 = FrameName[0];
            string temp_string2 = FrameName[0];
            //第一根线框
            ret = mySapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 10, ref temp_string1, "R1", "1", "Global");
            FrameName[0] = temp_string1;
            //第二根线框
            ret = mySapModel.FrameObj.AddByCoord(0, 0, 10, 8, 0, 16, ref temp_string1, "R1", "2", "Global");
            FrameName[1] = temp_string1;
            //第三根线框
            ret = mySapModel.FrameObj.AddByCoord(-4, 0, 10, 0, 0, 10, ref temp_string1, "R1", "3", "Global");
            FrameName[2] = temp_string1;

            //增加约束 PointName 两个约束点
            string[] PointName = new string[2];
            //定义下部的约束点
            //定义约束 1 2 3 4 约束 5 6 不约束 一个节点6个自由度
            bool[] Restraint = new bool[6];
            for (i = 0; i <= 3; i++)
            {
                Restraint[i] = true;
            }
            for (i = 4; i <= 5; i++)
            {
                Restraint[i] = false;
            }
            //获取第一根线的两个节点 temp_string1 temp_string2
            ret = mySapModel.FrameObj.GetPoints(FrameName[0], ref temp_string1, ref temp_string2);
            PointName[0] = temp_string1;
            PointName[1] = temp_string2;
            //将上面定义的约束 赋给第一个节点
            ret = mySapModel.PointObj.SetRestraint(PointName[0], ref Restraint, 0);

            //定义顶部的约束点 该约束点只约束 0 1 两个自由度 其他自由度放开
            for (i = 0; i <= 1; i++)
            {
                Restraint[i] = true;
            }
            for (i = 2; i <= 5; i++)
            {
                Restraint[i] = false;
            }
            //获取第二根线的两个节点 
            ret = mySapModel.FrameObj.GetPoints(FrameName[1], ref temp_string1, ref temp_string2);
            PointName[0] = temp_string1;
            PointName[1] = temp_string2;
            //将上面定义的约束赋给第二个节点
            ret = mySapModel.PointObj.SetRestraint(PointName[1], ref Restraint, 0);

            //刷新sap2000界面的视图
            bool temp_bool = false;
            ret = mySapModel.View.RefreshView(0, temp_bool);

            //增加7种载荷样式
            temp_bool = true;
            ret = mySapModel.LoadPatterns.Add("1", eLoadPatternType.Other, 1, temp_bool);
            ret = mySapModel.LoadPatterns.Add("2", eLoadPatternType.Other, 0, temp_bool);
            ret = mySapModel.LoadPatterns.Add("3", eLoadPatternType.Other, 0, temp_bool);
            ret = mySapModel.LoadPatterns.Add("4", eLoadPatternType.Other, 0, temp_bool);
            ret = mySapModel.LoadPatterns.Add("5", eLoadPatternType.Other, 0, temp_bool);
            ret = mySapModel.LoadPatterns.Add("6", eLoadPatternType.Other, 0, temp_bool);
            ret = mySapModel.LoadPatterns.Add("7", eLoadPatternType.Other, 0, temp_bool);

            //对第二种载荷样式添加载荷
            //获取第3根线的两个节点
            ret = mySapModel.FrameObj.GetPoints(FrameName[2], ref temp_string1, ref temp_string2);
            PointName[0] = temp_string1;
            PointName[1] = temp_string2;
            //定义载荷 第3个自由度-10 其他方向无载荷
            double[] PointLoadValue = new double[6];
            PointLoadValue[2] = -10;
            //第一个节点赋载荷
            ret = mySapModel.PointObj.SetLoadForce(PointName[0], "2", ref PointLoadValue, false, "Global", 0);
            //添加分布载荷
            ret = mySapModel.FrameObj.SetLoadDistributed(FrameName[2], "2", 1, 10, 0, 1, 1.8, 1.8, "Global", System.Convert.ToBoolean(-1), System.Convert.ToBoolean(-1), 0);

            //对第三种载荷样式添加载荷
            ret = mySapModel.FrameObj.GetPoints(FrameName[2], ref temp_string1, ref temp_string2);
            PointName[0] = temp_string1;
            PointName[1] = temp_string2;
            //定义载荷
            PointLoadValue = new double[6];
            PointLoadValue[2] = -17.2;
            PointLoadValue[4] = -54.4;
            //第二节节点赋载荷
            ret = mySapModel.PointObj.SetLoadForce(PointName[1], "3", ref PointLoadValue, false, "Global", 0);

            //载荷样式五添加载荷 第二根线添加分布载荷
            ret = mySapModel.FrameObj.SetLoadDistributed(FrameName[1], "4", 1, 11, 0, 1, 2, 2, "Global", System.Convert.ToBoolean(-1), System.Convert.ToBoolean(-1), 0);

            //样式五添加分布载荷
            ret = mySapModel.FrameObj.SetLoadDistributed(FrameName[0], "5", 1, 2, 0, 1, 2, 2, "Local", System.Convert.ToBoolean(-1), System.Convert.ToBoolean(-1), 0);
            ret = mySapModel.FrameObj.SetLoadDistributed(FrameName[1], "5", 1, 2, 0, 1, -2, -2, "Local", System.Convert.ToBoolean(-1), System.Convert.ToBoolean(-1), 0);

            //样式六添加分布载荷
            ret = mySapModel.FrameObj.SetLoadDistributed(FrameName[0], "6", 1, 2, 0, 1, 0.9984, 0.3744, "Local", System.Convert.ToBoolean(-1), System.Convert.ToBoolean(-1), 0);
            ret = mySapModel.FrameObj.SetLoadDistributed(FrameName[1], "6", 1, 2, 0, 1, -0.3744, 0, "Local", System.Convert.ToBoolean(-1), System.Convert.ToBoolean(-1), 0);

            //样式7添加点载荷
            ret = mySapModel.FrameObj.SetLoadPoint(FrameName[1], "7", 1, 2, 0.5, -15, "Local", System.Convert.ToBoolean(-1), System.Convert.ToBoolean(-1), 0);

            //设置单位制 kip_in_F
            ret = mySapModel.SetPresentUnits(eUnits.kip_in_F);

            //保存模型
            ret = mySapModel.File.Save(ModelPath);

            //分析模型 这一步会创建分析模型文件
            ret = mySapModel.Analyze.RunAnalysis();

            //下面的类容是后处理部分 暂时不用看 后面再研究
















            //initialize for SAP2000 results

            double[] SapResult = new double[7];

            ret = mySapModel.FrameObj.GetPoints(FrameName[1], ref temp_string1, ref temp_string2);

            PointName[0] = temp_string1;

            PointName[1] = temp_string2;



            //get SAP2000 results for load patterns 1 through 7           

            int NumberResults = 0;

            string[] Obj = new string[1];

            string[] Elm = new string[1];

            string[] LoadCase = new string[1];

            string[] StepType = new string[1];

            double[] StepNum = new double[1];

            double[] U1 = new double[1];

            double[] U2 = new double[1];

            double[] U3 = new double[1];

            double[] R1 = new double[1];

            double[] R2 = new double[1];

            double[] R3 = new double[1];

            for (i = 0; i <= 6; i++)

            {

                ret = mySapModel.Results.Setup.DeselectAllCasesAndCombosForOutput();

                ret = mySapModel.Results.Setup.SetCaseSelectedForOutput(System.Convert.ToString(i + 1), System.Convert.ToBoolean(-1));

                if (i <= 3)

                {

                    ret = mySapModel.Results.JointDispl(PointName[1], eItemTypeElm.ObjectElm, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);

                    U3.CopyTo(U3, 0);

                    SapResult[i] = U3[0];

                }

                else

                {

                    ret = mySapModel.Results.JointDispl(PointName[0], eItemTypeElm.ObjectElm, ref NumberResults, ref Obj, ref Elm, ref LoadCase, ref StepType, ref StepNum, ref U1, ref U2, ref U3, ref R1, ref R2, ref R3);

                    U1.CopyTo(U1, 0);

                    SapResult[i] = U1[0];

                }

            }



            //close SAP2000

            mySapObject.ApplicationExit(false);

            mySapModel = null;

            mySapObject = null;



            //fill SAP2000 result strings

            string[] SapResultString = new string[7];

            for (i = 0; i <= 6; i++)

            {

                SapResultString[i] = string.Format("{0:0.00000}", SapResult[i]);

                ret = (string.Compare(SapResultString[i], 1, "-", 1, 1, true));

                if (ret != 0)

                {

                    SapResultString[i] = " " + SapResultString[i];

                }

            }



            //fill independent results

            double[] IndResult = new double[7];

            string[] IndResultString = new string[7];

            IndResult[0] = -0.02639;

            IndResult[1] = 0.06296;

            IndResult[2] = 0.06296;

            IndResult[3] = -0.2963;

            IndResult[4] = 0.3125;

            IndResult[5] = 0.11556;

            IndResult[6] = 0.00651;

            for (i = 0; i <= 6; i++)

            {

                IndResultString[i] = string.Format("{0:0.00000}", IndResult[i]);

                ret = (string.Compare(IndResultString[i], 1, "-", 1, 1, true));

                if (ret != 0)

                {

                    IndResultString[i] = " " + IndResultString[i];

                }

            }



            //fill percent difference

            double[] PercentDiff = new double[7];

            string[] PercentDiffString = new string[7];

            for (i = 0; i <= 6; i++)

            {

                PercentDiff[i] = (SapResult[i] / IndResult[i]) - 1;

                PercentDiffString[i] = string.Format("{0:0%}", PercentDiff[i]);

                ret = (string.Compare(PercentDiffString[i], 1, "-", 1, 1, true));

                if (ret != 0)

                {

                    PercentDiffString[i] = " " + PercentDiffString[i];

                }

            }



            //display message box comparing results

            string msg = "";

            msg = msg + "LC  Sap2000  Independent  %Diff\r\n";

            for (i = 0; i <= 5; i++)

            {

                msg = msg + string.Format("{0:0}", i + 1) + "    " + SapResultString[i] + "   " + IndResultString[i] + "       " + PercentDiffString[i] + "\r\n";

            }



            msg = msg + string.Format("{0:0}", i + 1) + "    " + SapResultString[i] + "   " + IndResultString[i] + "       " + PercentDiffString[i];

            Console.WriteLine(msg);

            Console.ReadKey();




        }
    }
}

