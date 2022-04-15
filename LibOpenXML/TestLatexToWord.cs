using Latex2MathML;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LibOpenXML
{
    public class TestLatexToWord
    {
		LatexToMathMLConverter lmm;

		public void Main()
		{
			Convert();
			Console.ReadKey();
		}
		
		//mainxml.Load(nameMain);
		//LoadingUtils.LoadProject(mainxml,this);

		private string Convert(String latexExpression)
		{
			//String latexExpression = @"\begin{document} $\frac{\mathrm d_{1}}{\mathrm d x} \big{ k g(x \big)}$ \end{document}";
			//String latexExpression = @"\begin{document} $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$ \end{document}";
					

			lmm = new LatexToMathMLConverter();
			lmm.ValidateResult = true;
			//lmm.BeforeXmlFormat += MyEventListener;
			//lmm.ExceptionEvent += ExceptionListener;
			lmm.Convert(latexExpression);
			
			return lmm.Output;

		}

		private void Convert()
		{
			//return;
			//String latexExpression = @"\begin{document} $\frac{\mathrm d_{1}}{\mathrm d x} \big{ k g(x \big)}$ \end{document}";
			//String latexExpression = @"\begin{document} $x = \frac{-b \pm \sqrt{b^2 - 4ac}}{2a}$ \end{document}";
			//String latexExpression = @"\begin{document} $F_u=γ_t γ_c (RA+\sum_{}^{}\[R_(af,i) A_i \])$ \end{document}";
			String latexExpression = @"\begin{document} $F_u=γ_t γ_c \sum_{}^{} R_{(af,i)} A_{(af,i)}$ \end{document}";
			//"\[ \sum_{n=1}^{\infty} 2^{-n} = 1 \]"

			lmm = new LatexToMathMLConverter();
			lmm.ValidateResult = true;
			lmm.BeforeXmlFormat += MyEventListener;
			lmm.ExceptionEvent += ExceptionListener;
			lmm.Convert(latexExpression);

		}

		private void ExceptionListener(object sender, ExceptionEventArgs e)
		{
			Console.WriteLine("Exception handler called");
			String message = e.Message;
			Console.WriteLine(message);
		}

		private void MyEventListener(object sender, EventArgs e)
		{
			//Console.WriteLine("called .");
			String output = lmm.Output;
			Console.WriteLine(output);
			if (File.Exists("test.htm")) File.Delete("test.htm");
			File.AppendAllText("test.htm", output);
			System.Diagnostics.Process.Start("test.htm");
		}
	}
}
