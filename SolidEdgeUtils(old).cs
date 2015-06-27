using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

 using System.Windows.Forms;				// Per MessageBox()

using System.Runtime.InteropServices;		// Funzioni di base (COM)
using SolidEdgeFramework;					// Funzioni di base di Solid Edge
using SolidEdgeFrameworkSupport;
using System.Globalization;					// Per CultureInfo
using System.IO;							// Per DirectoryInfo

namespace Fred68.Tools.Grafica
	{
	class SolidEdgeUtils
		{
		class ProprietaSolidEdge				// Classe interna per archiviare le proprieta`
			{
			string mNome;						// Proprieta` pubbliche solo in lettura
			public string Nome
				{
				get { return mNome; }
				protected set { mNome = value; }
				}
			int mId;
			public int Id
				{
				get { return mId; }
				protected set { mId = value; }
				}
			dynamic mValore;					// Come tipo Variant ('var' non utilizzabile in una classe)
			public dynamic Valore				// In sola lettura
				{								// Scrittura concessa solo per alcuni tipi di dato (in costruttore o altri metodi)
				get { return mValore; }
				}
			public ProprietaSolidEdge(string nome, System.String valore, int id)
				{
				Nome = nome;
				Id = id;
				mValore = valore;
				}
			public ProprietaSolidEdge(string nome, System.Int32 valore, int id)
				{
				Nome = nome;
				Id = id;
				mValore = valore;
				}
			public ProprietaSolidEdge(string nome, System.Double valore, int id)
				{
				Nome = nome;
				Id = id;
				mValore = valore;
				}
			public ProprietaSolidEdge(string nome, System.DateTime valore, int id)
				{
				Nome = nome;
				Id = id;
				mValore = valore;
				}
			}
		// Variabili membro private legate a oggetti COM di Solid Edge 
		SolidEdgeFramework.Application application = null;			// Istanza
		Type type = null;											// Tipo di applicazione
		SolidEdgeFramework.SolidEdgeDocument document = null;		// Documento
		SEInstallDataLib.SEInstallData installData = null;			// Versione di Solid Edge...
			int builderNumber = 0;
			DirectoryInfo installFolder;
			CultureInfo cultureInfo;
			int majorVersion = 0;
			int minorVersion = 0;
			int parasolidMajorVersion = 0;
			int parasolidMinorVersion = 0;
			Version parasolidVersion;
			int servicePackVersion = 0;
			Version version;
		SolidEdgeDraft.ModelLinks modelLinks = null;				// Dati nel draft
		SolidEdgeDraft.ModelLink modelLink = null;
		
		SolidEdgeDraft.DrawingViews	drawingViews = null;			// Viste del foglio attivo
		SolidEdgeDraft.DrawingView drawingView = null;

		// SolidEdgeDraft.PartsLists partsLists = null;	...Distinte nel draft
		// SolidEdgeDraft.PartsList partsList = null;

		List<string> listaModelli = null;							// Lista dei modelli collegati al draft
		List<string> listaDocAperti = null;							// Lista dei documenti aperti
		List<string> listaProprietaDati = null;						// Nomi proprieta` con i 'dati'. La prima e` il codice di disegno
		List<ProprietaSolidEdge> listaProprietaFile = null;			// Lista delle proprieta` del file collegato
		List<string> listaScale = null;								// Lista delle scale

		string fullnameDoc, fullpdfDoc, fulldwgDoc;					// Nomi completi per salvataggio su file
		public bool nomecodicicorrispondono;

		public static string tipoDraft = "Draft";
		public static string tipoAssembly = "Assembly";
		public static string tipoPart = "Part";
		public static string tipoSheetMetal = "SheetMetal";
		public static string tipoSconosciuto = "Sconosciuto";
		public static string tipoWeldment = "Weldment";
		public static string tipoWeldmentAssembly = "WeldmentAssembly";
		public static string pathSeparator = "\\";					// Separatore di path
		public static string codeSeparator = ".";					// Separatore di codice es.: 780.36.220 (.)
		public static string revSeparator = "-";					// Separatore di revisione es.: 780.36.220-a (-)
		public static int lastLength = 3;							// Numero di caratteri dell'ultimo gruppo prima
																	// della revisione es.: 780.36.220a (3),
																	// usato se non c'e` separatore di revisione
		public static string altSeparator = "!";					// Separatore alternate assembly
		public static string pdfExt = ".pdf";						// Estensioni
		public static string dwgExt = ".dwg";
		public static string dftExt = ".dft";
		public static string dwgShExt = "-f";						// Estensione nomi fogli dwg

	//////////////////////////////////////////////////////
		public SolidEdgeUtils()										// Costruttore
			{
			listaModelli = new List<string>();
			listaDocAperti = new List<string>();
			listaProprietaDati = new List<string>();
			listaProprietaFile = new List<ProprietaSolidEdge>();
			listaScale = new List<string>();
			fullnameDoc = "";
			fullpdfDoc = "";
			fulldwgDoc = "";
			nomecodicicorrispondono = false;
			}
		public string AvviaSolidEdge()								// Avvia SE, restituisce string con messaggio
			{
			string messaggio = "";									// Legge eventuale istanza attiva. Eccezione se nessuna istanza
			try
				{
				application = (SolidEdgeFramework.Application)Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;									// Genera eccezione se SE non e` avviato
				}
			if (application != null)								// Se non e` attivo, lancia l'istanza
				{
				messaggio = "Istanza Solid Edge gia` in funzione";
				}
			else
				{
				try
					{																					// Prova avvio di Solid Edge
					type = Type.GetTypeFromProgID("SolidEdge.Application");								// Legge il tipo di appl. associato a SE
					application = (SolidEdgeFramework.Application)Activator.CreateInstance(type);		// Lancia un'istanza
					application.Visible = true;
					messaggio = "Avviato";
					}
				catch (System.Exception ex)							// Messaggio se eccezione
					{
					messaggio = "Eccezione: " + ex.Message;
					}
				finally												// Pulitura, eseguita sempre
					{
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return messaggio;
			}
		public string TipoDocumentoAttivo()							// Restituisce string con il tipo di documento attivo
			{
			string tipoDoc = "";				
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				tipoDoc = "Nessuna istanza attiva";
				application = null;
				}
			if(application != null)			// Se SE e` in funzione, cerca di leggere il tipo di documento attivo
				{
				try
					{
					document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
					switch (document.Type)
						{
						case DocumentTypeConstants.igAssemblyDocument:
							tipoDoc = tipoAssembly;
							break;
						case DocumentTypeConstants.igDraftDocument:
							tipoDoc = tipoDraft;
							break;
						case DocumentTypeConstants.igPartDocument:
							tipoDoc = tipoPart;
							break;
						case DocumentTypeConstants.igSheetMetalDocument:
							tipoDoc = tipoSheetMetal;
							break;
						case DocumentTypeConstants.igUnknownDocument:
							tipoDoc = tipoSconosciuto;
							break;
						case DocumentTypeConstants.igWeldmentAssemblyDocument:
							tipoDoc = tipoWeldmentAssembly;
							break;
						case DocumentTypeConstants.igWeldmentDocument:
							tipoDoc = tipoWeldment;
							break;
						}
					}
				catch
					{
					tipoDoc = "Nessun documento attivo";
					}
				finally
					{
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return tipoDoc;
			}
		public string NomeDocumentoAttivo()							// Restituisce il nome completo del doc attivo
			{
			string nomeDoc = "";
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;
				}
			if (application != null)			// Se SE e` in funzione, cerca di leggere il tipo di documento attivo
				{
				try
					{
					document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
					nomeDoc = document.FullName;
					}
				catch
					{
					nomeDoc = "";
					}
				finally
					{
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return nomeDoc;
			}
		public string VersioneSolidEdgeEInstallata()				// Restituisce string con i dati della versione
			{
			string versione = "";
			try
				{
				installData = new SEInstallDataLib.SEInstallData();
				builderNumber = installData.GetBuildNumber();
				installFolder = new DirectoryInfo(installData.GetInstalledPath());
				cultureInfo = new CultureInfo(installData.GetLanguageID());
				majorVersion = installData.GetMajorVersion();
				minorVersion = installData.GetMinorVersion();
				parasolidMajorVersion = installData.GetParasolidMajorVersion();
				parasolidMinorVersion = installData.GetParasolidMinorVersion();
				parasolidVersion = new Version(installData.GetParasolidVersion());
				servicePackVersion = installData.GetServicePackVersion();
				version = new Version(installData.GetVersion());
				versione = string.Format("Versione di Solid Edge :  {1}\nVersione parasolid :  {0}\nCartella di installazione :  {2}", parasolidVersion.ToString(), version.ToString(), installFolder.FullName);
				}
			catch (System.Exception ex)
				{
				Console.WriteLine(ex.Message);
				}
			finally
				{
				if (installData != null)
					{
					Marshal.ReleaseComObject(installData);
					installData = null;
					}
				}
			return versione;
			}
		public string CreaListaModelli()							// Riempie la lista dei nomi file collegati e rest. string messaggio
			{
			string messaggio = "";									// Cancella i dati preesistenti
			listaModelli.Clear();
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				messaggio = "Nessuna istanza attiva";
				application = null;
				}
			try
				{
				document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
				if(document.Type == DocumentTypeConstants.igDraftDocument)					// Se e` un draft...
					{
					int count, i;															// ...estrae i documenti linkati
					modelLinks = ((SolidEdgeDraft.DraftDocument)document).ModelLinks;
					count = modelLinks.Count;
					for (i = 0; i < count; i++)												// Li inserisce in una lista
						{
						modelLink = modelLinks.Item(i + 1);
						string nomeFile = modelLink.FileName;								// Filename completo
						int indx = nomeFile.IndexOf(altSeparator);							// Elimina alternate assembly
						if( indx != -1)
							{
							nomeFile = nomeFile.Remove(indx);								// Tutto cio` dopo '!'
							}
						listaModelli.Add(nomeFile);
						}
					}
				else
					{
					messaggio = "Non e` un draft";
					}
				}
			catch (System.Exception ex)							// Messaggio se eccezione
				{
				messaggio = "Eccezione: " + ex.Message + " + Nessun documento attivo";
				}
			finally
				{
				if (modelLinks != null)
					{
					foreach(SolidEdgeDraft.ModelLink modelLink in modelLinks)
						{
						Marshal.ReleaseComObject(modelLink);
						}
					}
				if (modelLinks != null)
					{
					Marshal.ReleaseComObject(modelLinks);
					modelLinks = null;
					}

				if (document != null)
					{
					Marshal.ReleaseComObject(document);
					document = null;
					}
				if (application != null)
					{
					Marshal.ReleaseComObject(application);
					application = null;
					}
				}
			return messaggio;
			}
		public List<string> ListaModelli							// Proprieta` read only
			{
			get {return listaModelli;}
			}
		public string CreaListaProprietaDaFile(string fullname)		// Lista da doc collegato su discio (poi resta in read only)
			{
			string msg = "";
			if(DocumentoAperto(fullname))
				{
				return("Documento aperto");
				}
			SolidEdgeFileProperties.PropertySets propertySets3D = null;	// Oggetti proprieta` del 3D
			SolidEdgeFileProperties.Properties properties3D = null;
			FileInfo dis3D = new FileInfo(fullname);				// Crea rif. al file
			if (dis3D.Exists)										// Se il file esiste
				{
				listaProprietaFile.Clear();									// Cancella la lista delle proprieta  attuale
				try
					{
					propertySets3D = new SolidEdgeFileProperties.PropertySets();					// Oggetto PropertySets
					propertySets3D.Open(fullname, true);											// Apre proprieta` file readonly (true)
					properties3D = (SolidEdgeFileProperties.Properties)propertySets3D["Custom"];	// Custom Property Set
					foreach (SolidEdgeFileProperties.Property pr in properties3D)					// Legge le proprieta`
						{
						listaProprietaFile.Add(new ProprietaSolidEdge(pr.Name, pr.Value, pr.ID));			// e le mette in una lista
						if (pr != null)
							{
							Marshal.ReleaseComObject(pr);
							}
						}					
					msg = "Lettura proprieta` completata";
					}
				catch (System.Exception ex)
					{
					msg += ex.Message + " in CreaListaProprietaDaFile()";
					}
				finally													// ATTENZIONE: NON RILASCIA IL FILE 3D, che rimane read-only !						
					{
					if (properties3D != null)
						{
						Marshal.ReleaseComObject(properties3D);
						properties3D = null;
						}
					if (propertySets3D != null)
						{
						propertySets3D.Close();							// Chiude il file prima di rilasciare il COM
						Marshal.ReleaseComObject(propertySets3D);
						Marshal.FinalReleaseComObject(propertySets3D);						
						propertySets3D = null;
						}
					GC.Collect();										// Per evitare errore di share violation dopo rilascio dei COM ?
					GC.WaitForPendingFinalizers();						// Ma il file resta aperto.
					}
				}	// fine blocco if(dis3D.Exists)
			else
				{
				msg = "File inesistente";
				}
			return msg;
			}
		public string CreaListaProprietaDaIstanzaSolidEdge(string fullname)	// Lista da doc collegato aperto
			{
			string msg = "";
			SolidEdgeFramework.SolidEdgeDocument activeDocument = null;		// Documento attivo
			SolidEdgeFramework.PropertySets propertySets = null;			// Proprieta` del documento aperto (non del file)
			SolidEdgeFramework.Properties properties = null;
			//SolidEdgeFramework.Property property = null;
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;
				}
			if (application != null)								// Se SE e` in funzione...
				{
				listaProprietaFile.Clear();							// Cancella la lista delle proprieta attuale
				try
					{
					activeDocument = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;	// Salva il rif. al doc. attivo
					if(AttivaDocumento(fullname))					// Attiva il documento aperto con il nome file richiesto...
						{
						msg += " - ";
						application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
						document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;	// Ne legge i rif. alle proprieta`
						propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
						properties = propertySets.Item("Custom");
						foreach (SolidEdgeFramework.Property pr in properties)							// Percorre le proprieta`
							{
							listaProprietaFile.Add(new ProprietaSolidEdge(pr.Name, pr.get_Value(), 0));				// e le mette in una lista
							if (pr != null)
								{
								Marshal.ReleaseComObject(pr);
								}
							}
						msg = "Lettura proprieta` completata";
						}			
					}
				catch (System.Exception ex)
					{
					msg += ex.Message + " in CreaListaProprietaDaIstanzaSolidEdge()";
					}
				finally
					{
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					activeDocument.Activate();					// Ripristina il doc che era attivo all'inizio della funzione
					}
				}
			return msg;
			}	
		public string VediListaProprieta()							// Crea una stringa con la lista delle proprieta`
			{
			string msg = string.Format("Trovate N. {0} proprieta` custom:\n",listaProprietaFile.Count);
			foreach (ProprietaSolidEdge pr in listaProprietaFile)					// Legge le proprieta`
				{
				msg += string.Format("ID = {0} : Name = {1} : Value = {2} : Type = {3}\n",
								pr.Id,
								pr.Nome,
								pr.Valore,
								pr.Valore.GetType());
				}					
			msg += "\nProprieta` filtrate:\n";
			foreach (ProprietaSolidEdge pr in listaProprietaFile)					// Legge le proprieta`
				{
				if(listaProprietaDati.Contains(pr.Nome))
					msg += string.Format("Name = {0} : Value = {1} : Type = {2}\n",
								pr.Nome,
								pr.Valore,
								pr.Valore.GetType());
				}					
			return msg;
			}
		public bool CreaListaDocumentiAperti()						// Crea lista documenti aperti 
			{
			bool ret = false;
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;
				}
			if (application != null)			// Se SE e` in funzione, leggere i documenti aperti e li mette in una lista
				{
				listaDocAperti.Clear();
				try
					{
					foreach(SolidEdgeFramework.SolidEdgeDocument doc in application.Documents)
						{
						listaDocAperti.Add(doc.FullName);
						}
					ret = true;
					}
				catch
					{
					ret = false;
					}
				finally
					{
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return ret;
			}
		public List<string> ListaDocAperti							// Proprieta` read only
			{
			get { return listaDocAperti; }
			}
		public bool DocumentoAperto(string fullname)				// Restituisce true se il doc indicato e` aperto nella sessione SE
			{
			bool ret = false;
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;
				}
			if (application != null)			// Se SE e` in funzione, leggere i documenti aperti e li mette in una lista
				{
				try
					{
					foreach (SolidEdgeFramework.SolidEdgeDocument doc in application.Documents)
						{
						if(doc.FullName == fullname)
							{
							ret = true;
							break;
							}
						}
					}
				catch
					{
					ret = false;
					}
				finally
					{
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return ret;
			}
		public bool AttivaDocumento(string fullname)				// Attiva il doc specificato (gia` aperto)
			{
			bool ret = false;
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;
				}
			if (application != null)			// Se SE e` in funzione, legge i documenti aperti e li mette in una lista
				{
				try
					{
					foreach (SolidEdgeFramework.SolidEdgeDocument doc in application.Documents)
						{
						if (doc.FullName == fullname)
							{
							doc.Activate();
							ret = true;
							break;
							}
						}
					}
				catch
					{
					ret = false;
					}
				finally
					{
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return ret;
			}	
		public bool ScriviListaProprieta(bool bTutte)				// Scrive lista prop nel doc attivo (in lista dati o tutte)
			{
			bool ret = false;
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;
				}
			if(application != null)
				{
				SolidEdgeFramework.PropertySets propertySets = null;			// Proprieta` del documento aperto (non del file)
				SolidEdgeFramework.Properties properties = null;
				SolidEdgeFramework.Property property = null;
				try
					{
					document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
					propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
					properties = propertySets.Item("Custom");
					foreach(ProprietaSolidEdge pr in listaProprietaFile)		// Percorre la lista proprieta`
						{
						if(listaProprietaDati.Contains(pr.Nome) || bTutte)		// Copia se in lista o se flag
							property = properties.Add(pr.Nome, pr.Valore);
						}
					ret = true;
					}
				catch								// Restituisce false se eccezione
					{
					ret = false;
					}
				finally
					{
					if (property != null)
						{
						Marshal.ReleaseComObject(property);
						property = null;
						}
					if (properties != null)
						{
						Marshal.ReleaseComObject(properties);
						properties = null;
						}
					if (propertySets != null)
						{
						Marshal.ReleaseComObject(propertySets);
						propertySets = null;
						}
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return ret;
			}
		public bool LeggiListaNomiDati(string fullname)				// Legge da file i nomi delle proprieta 'dati'
			{
			bool ret = false;
			FileInfo fileDati = new FileInfo(fullname);				// Rif. al file
			if(fileDati.Exists)										// Se il file esiste...
				{
				StreamReader fsr = fileDati.OpenText();				// Apre uno stream...
				List<string> lst = new List<string>();
				string line;
				while( (line = fsr.ReadLine()) != null)				// Legge tutte le righe
					{
					lst.Add(line);									// e le aggiunge alla lista
					}
				fsr.Close();										// Chiude lo stream
				if(lst.Count > 0)									// Se ha letto almeno una riga...
					{
					listaProprietaDati = lst;						// memorizza la lista al posto di quella vecchia
					ret = true;										// e imposta il flag
					}
				}
			return ret;			
			}
		public string CreaListaScale()								// Riempie una lista con le scale delle viste del draft
			{
			string messaggio = "";									// Cancella i dati preesistenti
			listaScale.Clear();
			try {
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				} catch {
				messaggio = "Nessuna istanza attiva";
				application = null;
				}
			try {
				document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
				if (document.Type == DocumentTypeConstants.igDraftDocument)					// Se e` un draft...
					{
					int count, i;															// ... estrae le viste del foglio attivo
					double scala;
					string tScala;
					drawingViews = ((SolidEdgeDraft.DraftDocument)document).ActiveSheet.DrawingViews;
					count = drawingViews.Count;
					for (i = 0; i < count; i++)												// Percorre le viste
						{
						drawingView = drawingViews.Item(i+1);
						scala = drawingView.ScaleFactor;
						if(scala >= 1.0)
							tScala = scala.ToString()+":1";
						else
							tScala = "1:"+(1.0/scala).ToString();
						listaScale.Add(tScala);
						}
					}
				else {
					messaggio = "Non e` un draft";
					}
				} catch (System.Exception ex)							// Messaggio se eccezione
				{
				messaggio = "Eccezione: " + ex.Message + " + Nessun documento attivo";
				}
			finally {
				if (drawingViews != null) {
					foreach (SolidEdgeDraft.DrawingView drawingView in drawingViews) {
						Marshal.ReleaseComObject(drawingView);
						}
					}
				if (drawingViews != null) {
					Marshal.ReleaseComObject(drawingViews);
					drawingViews = null;
					}

				if (document != null) {
					Marshal.ReleaseComObject(document);
					document = null;
					}
				if (application != null) {
					Marshal.ReleaseComObject(application);
					application = null;
					}
				}
			return messaggio;
			}
		public List<string> ListaScale								// Restituisce la lista scale calcolata 
			{
			get { return listaScale; }
			}
		public string NomeFoglioBackgroundAttivo()					// Restituisce il nome del foglio di background attivo 
			{
			string messaggio = "";									// Cancella i dati preesistenti
			try {
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				} catch {
				messaggio = "Nessuna istanza attiva";
				application = null;
				}
			try {
				document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
				if (document.Type == DocumentTypeConstants.igDraftDocument)					// Se e` un draft...
					{
					messaggio = ((SolidEdgeDraft.DraftDocument)document).ActiveSheet.Background.Name;
					}
				else {
					messaggio = "Non e` un draft";
					}
				} catch (System.Exception ex)							// Messaggio se eccezione
				{
				messaggio = "Eccezione: " + ex.Message + " + Nessun documento attivo";
				}
			finally {
				if (drawingViews != null) {
					foreach (SolidEdgeDraft.DrawingView drawingView in drawingViews) {
						Marshal.ReleaseComObject(drawingView);
						}
					}
				if (drawingViews != null) {
					Marshal.ReleaseComObject(drawingViews);
					drawingViews = null;
					}

				if (document != null) {
					Marshal.ReleaseComObject(document);
					document = null;
					}
				if (application != null) {
					Marshal.ReleaseComObject(application);
					application = null;
					}
				}
			return messaggio;
			}
		public bool ScriviSingolaProprieta(string nomeP, string valoreP)	// Scrive singola prop text nel doc attivo
			{
			bool ret = false;
			try {
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				} catch {
				application = null;
				}
			if (application != null) {
				SolidEdgeFramework.PropertySets propertySets = null;			// Proprieta` del documento aperto
				SolidEdgeFramework.Properties properties = null;
				SolidEdgeFramework.Property property = null;
				try {
					document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
					propertySets = (SolidEdgeFramework.PropertySets)document.Properties;
					properties = propertySets.Item("Custom");
					property = properties.Add(nomeP, valoreP);
					ret = true;
					}
				catch								// Restituisce false se eccezione
					{
					ret = false;
					}
				finally {
					if (property != null) {
						Marshal.ReleaseComObject(property);
						property = null;
						}
					if (properties != null) {
						Marshal.ReleaseComObject(properties);
						properties = null;
						}
					if (propertySets != null) {
						Marshal.ReleaseComObject(propertySets);
						propertySets = null;
						}
					if (document != null) {
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null) {
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return ret;
			}
		public string SaveDocumentoAttivo()							// Salva e restituisce il nome completo del doc attivo
			{
			string nomeDoc = "";
			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;
				}
			if (application != null)			// Se SE e` in funzione, cerca di leggere il tipo di documento attivo
				{
				try
					{
					document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;	// Di tipo Draft
					nomeDoc = document.FullName;
					document.Save();
					document.SaveAs(fullpdfDoc);			// Salva il documento come pdf
					//document.SaveAs(fulldwgDoc);			// Salva solo il foglio attivo come dwg

					SolidEdgeDraft.Sheet ActiveSheet;		// Memorizza il foglio attivo
					ActiveSheet = ((SolidEdgeDraft.DraftDocument)document).ActiveSheet;

					// Percorre i fogli e li salva singolarmente in dwg
					SolidEdgeDraft.Sections sections = null;
					SolidEdgeDraft.Section section = null;
					SolidEdgeDraft.SectionSheets sectionSheets = null;
					SolidEdgeDraft.Sheet sheet = null;

					sections = ((SolidEdgeDraft.DraftDocument)document).Sections;	// Le sezioni del documento dft
					section = sections.WorkingSection;					// Sezione con i fogli di lavoro (non background)
					sectionSheets = section.Sheets;						// Ottiene la lista dei fogli
					
					for (int j = 1; j <= sectionSheets.Count; j++)		// Percorre tutti i fogli
						{
						sheet = sectionSheets.Item(j);
						sheet.Activate();
						string dwgSuffix = "";							// Se solo 1 foglio, nessuna estensione
						if(sectionSheets.Count> 1)
							dwgSuffix = dwgShExt + j.ToString();
						document.SaveAs(fulldwgDoc + dwgSuffix + dwgExt);	// Salva il foglio attivo come dwg con estensione
						}
					ActiveSheet.Activate();		// Riattiva il foglio memorizzato

					}
				catch
					{
					//nomeDoc = "";
					}
				finally
					{
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return nomeDoc;
			}
		public bool PreparaNomifile()								// Imposta i nomi completi del file attivo...
			{															// ...solo se e` una dft
			bool done = false;
			string simplenameDoc = "";
			string pathDoc = "";
			string codenameDoc = "";
			fullnameDoc = "";
			fullpdfDoc = "";
			fulldwgDoc = "";											// Azzera i nomi

			try
				{
				application = (SolidEdgeFramework.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("SolidEdge.Application");
				}
			catch
				{
				application = null;
				}
			if (application != null)			// Se SE e` in funzione, cerca di leggere il tipo di documento attivo
				{
				try
					{
					document = (SolidEdgeFramework.SolidEdgeDocument)application.ActiveDocument;
					if(document.Type == DocumentTypeConstants.igDraftDocument)
						{
						fullnameDoc = document.FullName;								// Nome completo
						string codice = "";
						string[] split = document.Name.Split(new Char [] {'.'});						
						simplenameDoc = split[0];									// Nome senza estensioni
						pathDoc = document.Path + SolidEdgeUtils.pathSeparator;		// Path con separatore
						CreaListaProprietaDaIstanzaSolidEdge(fullnameDoc);			// Riempie lista proprieta` del doc attivo
						foreach (ProprietaSolidEdge pr in listaProprietaFile)		// Legge le proprieta`
							{
							if(pr.Nome == listaProprietaDati[0])					// Estrae quelle con il primo nome
								{													//  del file di configurazione
								codice = pr.Valore;
								}
							}
						List<string>[] lst = Codici(codice);						// Estrae codici ed estensioni
						foreach(string s in lst[0])									
							{
							codenameDoc += s;										// Nomi senza estensione
							}

						if(codenameDoc != simplenameDoc)							// Corrispondenza nomefile e codice							
							nomecodicicorrispondono = false;
						else
							nomecodicicorrispondono = true;

						foreach(string s in lst[1])									// Aggiunge estensione
							{
							codenameDoc += s;
							}
						fullpdfDoc = pathDoc + "PDF" + SolidEdgeUtils.pathSeparator + codenameDoc + pdfExt; 
						fulldwgDoc = pathDoc + "DWG" + SolidEdgeUtils.pathSeparator + codenameDoc; //+ dwgExt; 
						//MessageBox.Show(fullnameDoc + "\n" + fullpdfDoc + "\n" + fulldwgDoc);
						done = true;
						}
					else
						{
						//MessageBox.Show(document.Type.ToString());
						}	
					}
				catch /*(System.Exception ex)*/
					{
					fullnameDoc = "";
					//MessageBox.Show(ex.Message);
					}
				finally
					{
					if (document != null)
						{
						Marshal.ReleaseComObject(document);
						document = null;
						}
					if (application != null)
						{
						Marshal.ReleaseComObject(application);
						application = null;
						}
					}
				}
			return done;
			}
		public List<string>[] Codici(string codice)					// Separa le parti di nome e revisione dal codice
			{
			List<string>[] liste = new List<string>[2];
			liste[0] = new List<string>();
			liste[1] = new List<string>();
			int pos;
			string cd = codice;
			string tmp;
			while( (pos = cd.IndexOf(codeSeparator)) != -1)			// Cerca il separatore
				{
				tmp = cd.Substring(0,pos);							// Estrae la prima parte
				liste[0].Add(tmp);
				cd = cd.Substring(pos+1);							// e la taglia
				
				}													// Poi, nella restante stringa...
			if( (pos = cd.IndexOf(revSeparator)) != -1)				// ...cerca il separatore di revisione
				{
				tmp = cd.Substring(0,pos);							// Estrae la prima parte
				liste[0].Add(tmp);
				cd = cd.Substring(pos+1);
				liste[1].Add(cd);
				}
			else
				{													// Se non c'e` separatore
				tmp = cd.Substring(0,lastLength);					// estrae ultimo blocco di codice
				liste[0].Add(tmp);
				if(cd.Length > lastLength)
					{
					cd = cd.Substring(lastLength);
					liste[1].Add(cd);
					}
				else
					cd = "";
				}
			return liste;
			}
		}
	}
