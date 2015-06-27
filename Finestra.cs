using System;
using System.Collections.Generic;
using System.Text;
using System.Drawing;

using Fred68.Tools.Matematica;

namespace Fred68.Tools.Grafica
	{
	/// <summary> Finestra con window e viewport </summary>
	public class Finestra
		{
		#pragma warning disable 1591
		/// <summary>
		/// Tipo di scaling
		/// </summary>
		public enum TipoScala {Isotropico, Anisotropico};
		public Point p1;											// Viewport
		public Point p2;
		public Point2D P1;											// Finestra reale
		public Point2D P2;
		protected double fattoreZoom = 1.1;
		protected double fattorePan = 0.05;
		protected TipoScala scaling;

		double sx,sy;												// Scale x e y
		double dx, dy;												// Ampiezza viewport
		double dX, dY;												// Ampiezza finestra
		Point2D centro;												// centro finestra, usato per gli zoom
		static readonly Point2D panx = new Point2D(1, 0);			// Versori di appoggio per le direzioni
		static readonly Point2D pany = new Point2D(0, 1);
		
		#region PROPRIETA
		public double FattoreZoom
			{
			get { return fattoreZoom; }
			set { fattoreZoom = value; }
			}
		public double FattorePan
			{
			get { return fattorePan; }
			set { fattorePan = value; }
			}
		public TipoScala tipoScala
			{
			get {return scaling;}
			set	{
				scaling = value;
				RicalcolaFinestra();
				}
			}
		public double SCALAX
			{
			get {return sx;}
			}
		public double SCALAY
			{
			get {return sy;}
			}
		#pragma warning restore 1591
		#endregion
		#region COSTRUTTORI
		/// <summary>
		/// Costruttore
		/// </summary>
		public Finestra()
			{
			P1 = new Point2D();		// Non null, se no errore al primo ricalcolo
			P2 = new Point2D();
			scaling = TipoScala.Isotropico;
			}
		#endregion
		#region IMPOSTAZIONE
		/// <summary>
		/// Imposta Window
		/// </summary>
		/// <param name="pt1"></param>
		/// <param name="pt2"></param>
		/// <returns></returns>
		public bool Set(Point2D pt1, Point2D pt2)
			{
			P1 = pt1;
			P2 = pt2;
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Imposta Window
		/// </summary>
		/// <param name="x1"></param>
		/// <param name="y1"></param>
		/// <param name="x2"></param>
		/// <param name="y2"></param>
		/// <returns></returns>
		public bool Set(double x1, double y1, double x2, double y2)
			{
			P1 = new Point2D(x1, y1);
			P2 = new Point2D(x2, y2);
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Imposta Wiewport
		/// </summary>
		/// <param name="pt1"></param>
		/// <param name="pt2"></param>
		/// <returns></returns>
		public bool Set(Point pt1, Point pt2)
			{
			p1 = pt1;
			p2 = pt2;
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Imposta viewport
		/// </summary>
		/// <param name="x1"></param>
		/// <param name="y1"></param>
		/// <param name="x2"></param>
		/// <param name="y2"></param>
		/// <returns></returns>
		public bool Set(int x1, int y1, int x2, int y2)
			{
			p1 = new Point(x1, y1);
			p2 = new Point(x2, y2);
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Ricalcola parametri e finestra
		/// </summary>
		/// <returns></returns>
		protected bool RicalcolaFinestra()
			{
			bool ok = false;
			dX = P2.x - P1.x;				// Ampiezza finestra
			dY = P2.y - P1.y;
			dx = p2.X - p1.X;				// Ampiezza viewport
			dy = p2.Y - p1.Y;
			if( (System.Math.Abs(dX) > Point2D.Epsilon) && (System.Math.Abs(dY) > Point2D.Epsilon) )
				{
				double sxnew, synew, scom;							// Nuovi fattori di scala
				sxnew = dx / dX;
				synew = dy / dY;
				scom = Math.Min(Math.Abs(sxnew), Math.Abs(synew));

				if( (System.Math.Abs(sxnew) > Point2D.Epsilon) && (System.Math.Abs(synew) > Point2D.Epsilon) )
					{
					switch(scaling)
						{
						case TipoScala.Isotropico:
							{
							sx = Math.Sign(sxnew) * scom;				// Imposta fattore di scala comune
							sy = Math.Sign(synew) * scom;
							P1 = Get(p1);								// Ricalcola la window, con stesso aspetto della vieport
							P2 = Get(p2);
							break;
							}
						case TipoScala.Anisotropico:
							{
							sx = sxnew;
							sy = synew;
							break;
							}
						}
					ok = true;
					}
				}
			centro = (P1+P2)/2;
			return ok;
			}
		/// <summary>
		/// Zoom out
		/// </summary>
		/// <returns></returns>
		public bool ZoomOut()
			{
			if(fattoreZoom <= 1.0)							// Se fattore errato, esce
				return false;
			centro = (P1 + P2) / 2.0;						// Punto centrale
			P1 = centro + (P1 - centro) * fattoreZoom;
			P2 = centro + (P2 - centro) * fattoreZoom;
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Zoom in
		/// </summary>
		/// <returns></returns>
		public bool ZoomIn()
			{
			if (fattoreZoom <= 1.0)						// Se fattore errato, esce
				return false;
			centro = (P1 + P2) / 2.0;						// Punto centrale
			P1 = centro + (P1 - centro) / fattoreZoom;
			P2 = centro + (P2 - centro) / fattoreZoom;
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Pan a destra
		/// </summary>
		/// <returns></returns>
		public bool PanDx()
			{
			if (fattorePan <= 0.0)						// Se fattore errato, esce
				return false;
			P1 = P1 + panx * dX * fattorePan;
			P2 = P2 + panx * dX * fattorePan;
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Pan a sinistra
		/// </summary>
		/// <returns></returns>
		public bool PanSx()
			{
			if (fattorePan <= 0.0)						// Se fattore errato, esce
				return false;
			P1 = P1 - panx * dX * fattorePan;
			P2 = P2 - panx * dX * fattorePan;
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Pan in su
		/// </summary>
		/// <returns></returns>
		public bool PanSu()
			{
			if (fattorePan <= 0.0)						// Se fattore errato, esce
				return false;
			P1 = P1 + pany * dY * fattorePan;
			P2 = P2 + pany * dY * fattorePan;
			RicalcolaFinestra();
			return true;
			}
		/// <summary>
		/// Pan in giu`
		/// </summary>
		/// <returns></returns>
		public bool PanGiu()
			{
			if (fattorePan <= 0.0)						// Se fattore errato, esce
				return false;
			P1 = P1 - pany * dY * fattorePan;
			P2 = P2 - pany * dY * fattorePan;
			RicalcolaFinestra();
			return true;
			}
		#endregion
		#region TRASFORMAZIONE COORDINATE
		/// <summary>
		/// Ottiene il punto corrispondente della viewport
		/// </summary>
		/// <param name="P">Punto della Window</param>
		/// <returns></returns>
		public Point Get(Point2D P)
			{
			return new Point( (int)((P.x - P1.x) * sx + p1.X), (int)((P.y - P1.y) * sy + p1.Y) );
			}
		/// <summary>
		/// Ottiene il punto corrispondente della Window
		/// </summary>
		/// <param name="p">Punto della Viewport</param>
		/// <returns></returns>
		public Point2D Get(Point p)
			{
			return new Point2D(  ((p.X - p1.X)/sx) + P1.x, ((p.Y - p1.Y)/sy) + P1.y );
			}
		#endregion
		}
	}
