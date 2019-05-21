/*Modify List:
 * 20050725,Color&Font add
 * 
 */
//need:
//C:\WINDOWS\Microsoft.NET\Framework\v1.1.4322\Microsoft.JScript.dll
using System;
using System.Drawing;
using System.Text;
using System.IO;
using System.Data;
using System.Collections;
using System.Collections.Specialized;
using System.Windows.Forms;

	/// <summary>
	/// Common Class
	/// </summary>
namespace cc
{
	/// <summary>
	/// Eval
	/// </summary>
	public class Eval
	{
		static Microsoft.JScript.Vsa.VsaEngine EvaluateToString_VsaEngine = Microsoft.JScript.Vsa.VsaEngine.CreateEngine();
		/// <summary>
		/// Summary description for Class:EvaluateToBool
		/// need:
		/// Microsoft.JScript.dll
		/// Microsoft.Vsa.dll
		/// </summary>
		//Sample1
		//string expr = @"
		//'1+2+3*4=' + (1+2+3*4);
		//return:1+2+3*4=15
		//
		//Sample2
		//string expr = @"
		//var x = 11;
		//var y = 'A';
		//var bEval = '1';
		//if((x > 10) && (y == 'A')){
		//	bEval = '2';
		//}
		//else{
		//	bEval = '3';
		//}";
		//return:2
		//
		//Sample3
		//string expr = @"
		//var x = 11;
		//var y = 'A';
		//x=x;}";
		//return:11   <-- the last value is return!
		static public object JSEvaluateToObject(string JScriptEvaluateString)
		{
			string ErrorString;
			return JSEvaluateToObject(JScriptEvaluateString, true, out ErrorString);
		}
		static public object JSEvaluateToObject(string JScriptEvaluateString, bool isEngineReset)
		{
			string ErrorString;
			return JSEvaluateToObject(JScriptEvaluateString, isEngineReset, out ErrorString);
		}
		static public object JSEvaluateToObject(string JScriptEvaluateString, bool isEngineReset, out string ErrorString)
		{
			ErrorString = null;
			try
			{
				if(isEngineReset)
				{
					EvaluateToString_VsaEngine.Reset();
				}
				return Microsoft.JScript.Eval.JScriptEvaluate(JScriptEvaluateString, EvaluateToString_VsaEngine);
			}
			catch(Exception exp)
			{
				ErrorString = "Run JScript Error:" + exp.Message + "\r\nSource:" + JScriptEvaluateString;
				return null;
			}
		}

		static DataTable SQLEvaluateToString_dt = new DataTable();
		/// <summary>
		/// Summary description for Class:EvaluateToBool
		/// </summary>
		static public object SQLEvaluateToObject(string sIF, string sOUT)
		{
			string ErrorString;
			return SQLEvaluateToObject(sIF, sOUT, out ErrorString);
		}
		static public object SQLEvaluateToObject(string sIF, string sOUT, out string ErrorString)
		{
			ErrorString = null;
			try
			{
				return SQLEvaluateToString_dt.Compute(sOUT, sIF);
			}
			catch(Exception exp)
			{
				ErrorString = "SQLEvaluateToString Eror:" + exp.Message + "\r\nIF:" + sIF + "\r\nOUT:" + sOUT;
				return null;
			}
		}
		
		static public object CSharpEvaluateToObject(string cCharpEvaluateString)
		{
			string ErrorString;
			return CSharpEvaluateToObject(cCharpEvaluateString, out ErrorString);
		}
		/// <summary>
		/// Summary description for Class:cCharpEvaluateString
		/// </summary>
		static public object CSharpEvaluateToObject(string cCharpEvaluateString, out string ErrorString)
		{
			ErrorString = "";
			try
			{
				Microsoft.CSharp.CSharpCodeProvider csharpCodeProvider = new Microsoft.CSharp.CSharpCodeProvider();
				System.CodeDom.Compiler.ICodeCompiler compiler = csharpCodeProvider.CreateCompiler();
				System.CodeDom.Compiler.CompilerParameters cp = new System.CodeDom.Compiler.CompilerParameters();
				cp.ReferencedAssemblies.Add("system.dll");
				cp.CompilerOptions = "/t:library";
				cp.GenerateInMemory = true;

				StringBuilder sCode = new StringBuilder();
				sCode.Append("using System;");
				sCode.Append("namespace CoustomEvaluate{");
				sCode.Append("class A{public object Eval(){return (" + cCharpEvaluateString + ");}}");
				sCode.Append("}");

				System.CodeDom.Compiler.CompilerResults results = compiler.CompileAssemblyFromSource(cp, sCode.ToString());
				if(results.Errors.HasErrors)
				{
					ErrorString = "Compiler Error:\r\n";
					foreach(System.CodeDom.Compiler.CompilerError error in results.Errors)
					{
						ErrorString += "Error Line:" + error.Line + ", Column:" + error.Column + "\r\n";
						ErrorString += "ErrorText:" + error.ErrorText + "\r\n";
					}
					return null;
				}
				System.Reflection.Assembly assembly = results.CompiledAssembly;
				object Inst = assembly.CreateInstance("CoustomEvaluate.A");
				Type type = Inst.GetType();
				System.Reflection.MethodInfo mi = type.GetMethod("Eval");
				return mi.Invoke(Inst, null);
			}
			catch(Exception exp)
			{
				ErrorString = "CSharpEvaluateToObject Eror:" + exp.Message + "\r\nSource:" + cCharpEvaluateString;
				return null;
			}
		}
	}

}
