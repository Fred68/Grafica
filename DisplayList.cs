using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;
using Fred68.Tools.Matematica;
using Fred68.Tools.Grafica;


namespace Fred68.Tools.Grafica
	{
	/// <summary> Elemento della display list </summary>
	public class DisplayListElement 
		{
		#pragma warning disable 1591		
		protected Tratto tratto;
		protected int penna;
		public Tratto Tratto
			{
			get {return tratto;}
			}
		#pragma warning restore 1591
		/// <summary>
		/// Costruttore
		/// </summary>
		/// <param name="tr">Tratto</param>
		/// <param name="pen">Indice della penna</param>
		public DisplayListElement(Tratto tr, int pen)
			{
			tratto = tr;
			penna = pen;
			}
		/// <summary>
		/// Plot
		/// </summary>
		/// <param name="dc"></param>
		/// <param name="fin"></param>
		/// <param name="penne"></param>
		public void Plot(Graphics dc, Finestra fin, Pen[] penne)
			{
			if((penna >= penne.Length) || (penna<0))	return;
			tratto.Plot(dc,fin,penne[penna]);
			}
		}
	/// <summary> Display List </summary>
	public class DisplayList 
		{
		#warning Aggiungere funzione di calcolo dei punti minimo e massimo, da usarsi per FitZoom
		/// <summary>
		/// Lista degli elementi
		/// </summary>
		protected List<DisplayListElement> dl;
		/// <summary>
		/// Costruttore
		/// </summary>
		public DisplayList()
			{
			dl = new List<DisplayListElement>();
			dl.Clear();
			}
		/// <summary>
		/// Svuota
		/// </summary>
		public void Clear()
			{
			dl.Clear();
			}
		/// <summary>
		/// Aggiunge un elemento
		/// </summary>
		/// <param name="dle"></param>
		public void Add(DisplayListElement dle)
			{
			dl.Add(dle);
			}
		/// <summary>
		/// Plot
		/// </summary>
		/// <param name="dc"></param>
		/// <param name="fin"></param>
		/// <param name="penne"></param>
		public void Plot(Graphics dc, Finestra fin, Pen[] penne)
		    {
		    foreach(DisplayListElement dle in dl)
		        {
		        dle.Plot(dc, fin, penne);
		        }
		    }
		}
	}
