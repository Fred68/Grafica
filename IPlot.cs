using System.Drawing;

namespace Fred68.Tools.Grafica
	{
	interface IPlot
		{
		void Plot(Graphics dc, Finestra fin, Pen penna);
		void Display(DisplayList displaylist, int penna);
		}
	}