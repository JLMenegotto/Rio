using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Xml.Xsl;
using System.Xml.Linq;
using System.Reflection;

using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Colors;

using OX = Microsoft.Office.Interop.Excel;

namespace Rio
{
     public class Basicas       
	 {
		private Document    _doc;
		private Editor      _edi;
		private Database    _db1;
		private Database    _db2;
		private Transaction _tran;
 
		public Document    doc  { get { return Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument; }                         set { _doc = value; } }
		public Editor      edi  { get { return Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor; }                  set { _edi = value; } }
		public Database    dba  { get { return Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.Document.Database;} set { _db1 = value; } }
		public Database    dbw  { get { return HostApplicationServices.WorkingDatabase; }                                                                         set { _db2 = value; } }
		public Transaction tran { get { return _tran; }                                                                                                           set { _tran = value; } }

		public double       Str_Doble (string s) { return System.Convert.ToDouble(s);}
		public long         Str_Int64 (string s) { return System.Convert.ToInt64(s); }
		public int          Str_Int32 (string s) { return System.Convert.ToInt32(s); }

        public void         Mensaje   (string m)
        { 
                            System.Windows.Forms.MessageBox.Show(m);
        }
        public Point3d      PontoMedio (Point3d p1, Point3d p2)            
		{
			double dis = p1.DistanceTo(p2);
			return new Point3d((p1.X + p2.X) * 0.5, (p1.Y + p2.Y) * 0.5, (p1.Z + p2.Z) * 0.5);
		}
        public Point3d      PontoMedio (Point3d p1, Point3d p2, double fat)
		{
			double dis = p1.DistanceTo(p2);
			return new Point3d((p1.X + p2.X) * fat, (p1.Y + p2.Y) * fat, (p1.Z + p2.Z) * fat);
		}
        public double       Comprimento(Object curva)                      
		{
			Curve elem = curva as Curve;
			double comprimento = 0.0;
			if (elem != null)
			{
				comprimento = elem.GetDistanceAtParameter(elem.EndParam)
							- elem.GetDistanceAtParameter(elem.StartParam);
			}
			return comprimento;
		}

		public ObjectId[]   Captura_entidades(string entidade, bool msg )  
		{
			ObjectId[]  objid = null;
			SelectionSet sset = null;
			PromptSelectionResult res = null;
			try
			{
				TypedValue[] filtro = new TypedValue[] {
					                                     new TypedValue((int)DxfCode.Start,     entidade), 
					                                     new TypedValue((int)DxfCode.LayerName, "0") 
				                                       };

				SelectionFilter ss = new SelectionFilter(filtro);
				res   = edi.SelectAll(ss);
				sset  = res.Value;
				objid = sset.GetObjectIds();
				if (msg == true) Mensaje("Entidades capturada tipo " + entidade + ". Total = " + objid.Count().ToString());
				return objid;
			}
			catch { return objid; }
		}
		public ObjectId[]   Captura_Zonas    (string entidade, bool msg )  
		{
			ObjectId[] objid = null;
			SelectionSet sset = null;
			PromptSelectionResult res = null;
			try
			{
				TypedValue[] filtro = new TypedValue[] {
					                                     new TypedValue((int)DxfCode.Start,     entidade), 
					                                     new TypedValue((int)DxfCode.Operator,  "<or"), 
														 new TypedValue((int)DxfCode.LayerName, "Piano"), 
														 new TypedValue((int)DxfCode.LayerName, "Violin"),
														 new TypedValue((int)DxfCode.LayerName, "Tuba"),
														 new TypedValue((int)DxfCode.LayerName, "Fundo"),
														 new TypedValue((int)DxfCode.Operator,  "or>") 
				                                       };

				SelectionFilter ss = new SelectionFilter(filtro);
				res  = edi.SelectAll(ss);
				sset = res.Value;
				objid = sset.GetObjectIds();
				if (msg == true)
					Mensaje("Entidades capturada tipo " + entidade + ". Total = " + objid.Count().ToString());

				return objid;
			}
			catch { return objid; }
		}
		public Entity       Captura_Ultima   ()                            
		{
                            Entity ent;
                            ObjectId objid = Autodesk.AutoCAD.Internal.Utils.EntLast();
                            if (!objid.IsNull && objid.IsValid)
                               {
                                       ent = (Entity)objid.GetObject(OpenMode.ForWrite) as Entity;
                               }
                               else
                               {
                                       ent = null;
                               }
                     return ent;
		}

		public List<double> Caixa(Entity ent, out double dx, out double dy, out double dz)
		{
                            List<double> bbox = new List<double> { };
                            Extents3d bb = ent.GeometricExtents;
                            Point3d   p1 = bb.MinPoint;
                            Point3d   p2 = bb.MaxPoint;
                            Point3d   p0 = PontoMedio(p1, p2);
                            Vector3d  v1 = p1.GetAsVector();
                            Vector3d  v2 = p2.GetAsVector();
                            double  ang  = v1.GetAngleTo(v2) * (180 / Math.PI);
                            double  seno = System.Math.Sin(ang);
                            double  cose = System.Math.Cos(ang);
                                      dx = p1.DistanceTo(p2) * cose; bbox.Add(dx);
                                      dy = p1.DistanceTo(p2) * seno; bbox.Add(dy);
                                      dz = 0.0; bbox.Add(dz);
                     return bbox;
		}

		// ----------------------------------------------------------------------------------------------------------------
		public DBObject Cria_Esfera  (Transaction tr, Point3d centro, double raio, string nomelayer ) 
		{
			Matrix3d mat = new Matrix3d();
			Solid3d esf = new Solid3d();
			mat = Matrix3d.Displacement(centro.GetAsVector());
			NovaEntidad(tr, esf);
			esf.CreateSphere(raio);
			esf.TransformBy(mat);
			DBObject obj = tr.GetObject(esf.ObjectId, OpenMode.ForRead);
			esf.Layer = nomelayer;
			return obj;
		}
		public DBObject Cria_Circulo (Transaction tr, Point3d centro, double raio )                   
		{
			Circle cir = new Circle();
			cir.Center = centro;
			cir.Radius = raio;
			cir.Normal = Vector3d.ZAxis;
			NovaEntidad(tr, cir);
			DBObject obj = tr.GetObject(cir.ObjectId, OpenMode.ForRead);
			return obj;
		}
        public DBObject Cria_Circulo (Transaction tr, Point3d centro, double raio, string nomelayer ) 
		{
			Circle cir = new Circle();
			cir.Center = centro;
			cir.Radius = raio;
			cir.Normal = Vector3d.ZAxis;
			NovaEntidad(tr, cir);
			DBObject obj = tr.GetObject(cir.ObjectId, OpenMode.ForRead);
                  cir.Layer = nomelayer;
			return obj;
		}
		public DBObject Cria_Texto   (Transaction tr, Point3d pto, string txt )                       
		{
			Autodesk.AutoCAD.DatabaseServices.DBText tex = new Autodesk.AutoCAD.DatabaseServices.DBText();
			tex.TextStyleId = dba.Textstyle;
			tex.TextString  = txt;
			tex.Position    = pto;
			tex.Rotation    = 0.0;
			tex.Height      = 0.1;
			tex.Normal      = Vector3d.ZAxis;
			NovaEntidad(tr, tex);
			DBObject obj = tr.GetObject(tex.ObjectId, OpenMode.ForRead);
			edi.UpdateScreen();
			return obj;
		}
		public DBObject Cria_Hachura (Transaction tr, DBObject obj, double elev )                     
		{
			Hatch Hachura = new Hatch();
			NovaEntidad(tr, Hachura);

			ObjectIdCollection objcol = new ObjectIdCollection();
			objcol.Add(obj.ObjectId);

			Hachura.SetDatabaseDefaults();
			Hachura.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
			Hachura.AppendLoop(HatchLoopTypes.Default, objcol);

			Hachura.Color = Autodesk.AutoCAD.Colors.Color.FromRgb((byte)100, (byte)100, (byte)100);
			Hachura.Transparency = new Autodesk.AutoCAD.Colors.Transparency((byte)70);
			Hachura.Elevation = elev;
			Hachura.Associative = true;
			edi.UpdateScreen();
			return tr.GetObject(Hachura.ObjectId, OpenMode.ForRead);
		}
		public Entity   NovaEntidad  (Transaction tr, Entity  ent )                                   
		{
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(dba.CurrentSpaceId, OpenMode.ForWrite);
                        ObjectId obj = btr.AppendEntity(ent);
                        tr.AddNewlyCreatedDBObject(ent, true);
                        return ent;
		}

        public void         Faz_Offset (Polyline pl, Transaction tr, BlockTableRecord btr)            
        {
                    foreach (Entity ent in pl.GetOffsetCurves(0.025))
                    {
                             btr.AppendEntity(ent);
                             tr.AddNewlyCreatedDBObject(ent, true);
                    }
        }

		public ObjectId[]   Prepara_Zonas   ( )                                       
		{
			ObjectId[] PoligonaisZonas = Captura_Zonas("LWPOLYLINE", true);
			return PoligonaisZonas;
		}
		public string       Ponto_Dentro    ( Polyline poli, Point3d p0 )             
		{
			string puntoadentro = "-";
			Point3d p1 = poli.GetPoint3dAt(0);
			Point3d p2 = poli.GetPoint3dAt(1);
			Point3d p3 = poli.GetPoint3dAt(2);
			Point3d p4 = poli.GetPoint3dAt(3);

			double areapoli = poli.Area;
			string layerpol = poli.Layer;

			double areatri1 = AreaTriangulo(p0, p1, p2);
			double areatri2 = AreaTriangulo(p0, p2, p3);
			double areatri3 = AreaTriangulo(p0, p3, p4);
			double areatri4 = AreaTriangulo(p0, p4, p1);
			double somatori = areatri1 + areatri2 + areatri3 + areatri4;

			if (areapoli > (somatori * 0.995) && areapoli < (somatori * 1.005))
				 puntoadentro = layerpol;
			else
				 puntoadentro = "-";

			return puntoadentro;
		}
		public double       AreaTriangulo   ( Point3d p1,    Point3d p2, Point3d p3 ) 
		{
			double a = p1.DistanceTo(p2);
			double b = p2.DistanceTo(p3);
			double c = p3.DistanceTo(p1);
			double S = (a + b + c) / 2.0;
			double area = Math.Sqrt(S * (S - a) * (S - b) * (S - c));
			return area;
		}
		public Point3d      Seleccion_Ponto ( string msg )                            
		{
			PromptPointOptions pop = new PromptPointOptions("\n" + msg);
			PromptPointResult pnt = edi.GetPoint(pop);
			if (pnt.Status == PromptStatus.OK)
				return pnt.Value;
			else
				return new Point3d();
		}
             
		public int          Ingressar_Inteiro( string msg, int defaul)                
		{
			PromptIntegerResult numint;
			PromptIntegerOptions inte = new PromptIntegerOptions("\n" + msg);
			inte.DefaultValue = defaul;
			inte.UseDefaultValue = true;
			numint = edi.GetInteger(inte);
			return numint.Value;
		}
		public double       Ingressar_Real   ( string msg, double defaul )            
		{
			PromptDoubleResult numreal;
			PromptDoubleOptions real = new PromptDoubleOptions("\n" + msg);
			real.DefaultValue = defaul;
			real.UseDefaultValue = true;
			numreal = edi.GetDouble(real);
			return numreal.Value;
		}
		public string       Ingressar_Text   ( string msg, string defaul )            
		{
			PromptResult texto;
			PromptStringOptions textoin = new PromptStringOptions("\n" + msg);
			textoin.DefaultValue = defaul;
			textoin.UseDefaultValue = true;
			texto = edi.GetString(textoin);

			if (texto.Status != PromptStatus.OK)
				return defaul;
			else
				return texto.StringResult;
		}

		public bool         ParImp(int num)          { return num % 2   == 0; }           // Verifica se número é par ou impar       
		public bool         Modulo(int num, int mod) { return num % mod == 0; }           // Verifica indice modular do número       
		public int          Congru(int num, int mod)                                  
		{
                            int partes = (num / mod);
                            int modular = (num - (partes * mod));
                            return modular;
		} // Retorna congruente modular do número 

		public Entity       Selecionar_Entidad (string msg )                          
		{
			DBObject obj;
			Entity ent;
			PromptEntityResult per = edi.GetEntity("\n" + msg);
			Transaction tra = doc.TransactionManager.StartTransaction();
			using (tra)
			{
				obj = tra.GetObject(per.ObjectId, OpenMode.ForRead);
				ent = obj as Entity;
				return ent;
			}
		}	
		public SelectionSet FiltrarLayer (string layer)                               
		{
			    TypedValue[]          tvs  = new TypedValue[] { new TypedValue((int)DxfCode.LayerName, layer), };
			    SelectionFilter       sset = new SelectionFilter(tvs);
			    PromptSelectionResult psr  = edi.SelectAll( sset );
	            return psr.Value;
		}
		public ObjectId[]   FiltraObjetos(string entidade, string layer1)             
		{
			ObjectId[] objetos = null;
			SelectionSet sset  = null;
			PromptSelectionResult res = null;
			try
			{
				TypedValue[] filtro = new TypedValue[]  { 
                                                              new TypedValue((int)DxfCode.Operator,"<and"),
                                                              new TypedValue((int)DxfCode.Start,     entidade),
                                                              new TypedValue((int)DxfCode.LayerName, layer1),
                                                              new TypedValue((int)DxfCode.Operator,"and>")
                                                        };
				SelectionFilter ssfil = new SelectionFilter(filtro);
				res = edi.SelectAll(ssfil);
				sset = res.Value;
				objetos = sset.GetObjectIds();
				return objetos;
			}
			catch { return objetos; }
		}
		public ObjectId[]   Filtrar_Temas(string[] layer1)                            
		{
			    ObjectId[] objetos        = null;
			    SelectionSet sset         = null;
			    PromptSelectionResult res = null;
		         	string layers = "";
		      	foreach (string s in layer1)
			      {
			              layers = layers + "," + s;
			      }
  
			      try
			      {
				        TypedValue[] filtro = new TypedValue[]  { 
                                                                           new TypedValue((int)DxfCode.Operator,"<and"),
                                                                           new TypedValue(8, layers),
                                                                           new TypedValue((int)DxfCode.Operator,"and>")
                                                                       };
				    SelectionFilter ssfil = new SelectionFilter(filtro);
				    res = edi.SelectAll(ssfil);
				    sset = res.Value;
				    objetos = sset.GetObjectIds();
				return objetos;
			}
			catch { return objetos; }
		}
		
		public string Extrair_Coordenadas            (Entity enti, Transaction tr)    
		{
                          string t = enti.GetType().ToString();
                          string p = "-";
                          switch (t)
                          {
                                     case "Autodesk.AutoCAD.DatabaseServices.DBPoint":        p = Extrair_pontos_pontos(enti);               break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Line":           p = Extrair_pontos_linhas(enti);               break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Circle":         p = Extrair_pontos_circle(enti);               break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Polyline":       p = Extrair_pontos_poligonal2D(enti, 0.1);     break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Polyline2d":     p = Extrair_pontos_poligonal2D(enti, 0.1);     break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Polyline3d":     p = Extrair_pontos_poligonal3D(enti, tr, 0.1); break;
                                     case "Autodesk.AutoCAD.DatabaseServices.BlockReference": p = Extrair_pontos_bloco(enti);                break;
                                     case "Autodesk.AutoCAD.DatabaseServices.DBText":         p = Extrair_Pontos_texto(enti);                break;
                                     default:                                                 p = "-";                                       break;
                          }
                          return p;
		}  
        public string Tipo_de_Objeto                 (Entity enti)                    
		{ 
			              string  tipo  = enti.GetType().ToString();
                          string  obje  = "Desconocido";
			              switch (tipo)
			              {
                                     case "Autodesk.AutoCAD.DatabaseServices.DBPoint":        obje = "Ponto";       break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Line":           obje = "Linha";       break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Circle":         obje = "Circulo";     break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Polyline":       obje = "Poli_2D";     break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Polyline2d":     obje = "Poli_2D";     break;
                                     case "Autodesk.AutoCAD.DatabaseServices.Polyline3d":     obje = "Poli_3D";     break;
                                     case "Autodesk.AutoCAD.DatabaseServices.BlockReference": obje = "Bloco";       break;
                                     case "Autodesk.AutoCAD.DatabaseServices.DBText":         obje = "Texto";       break;
                                     default:                                                 obje = "Desconocido"; break;
                          }
                          return obje;
		}
        public string Texto_do_Objeto                (Entity enti)                    
		{
			              string tipo  = enti.GetType().ToString();
                          string texto = "-";
			              switch (tipo)
			              {
                                     case "Autodesk.AutoCAD.DatabaseServices.DBText":         texto = Extrair_texto(enti); break;
                                     default:                                                 texto = "-";                 break;
                          }
                          return texto;
		}
        public double Area_objeto                    (Entity enti)                    
        {
              double area = 0.0;
              Curve ob = enti as Curve;
              string tipo = enti.GetType().ToString();
			  switch (tipo)
			         {
                            case "Autodesk.AutoCAD.DatabaseServices.DBPoint":    break;
                            case "Autodesk.AutoCAD.DatabaseServices.Line":       break;
                            case "Autodesk.AutoCAD.DatabaseServices.Circle":     area = ob.Area; break;
                            case "Autodesk.AutoCAD.DatabaseServices.Polyline":   area = ob.Area; break;
                            case "Autodesk.AutoCAD.DatabaseServices.Polyline2d": area = ob.Area; break;
                            case "Autodesk.AutoCAD.DatabaseServices.Polyline3d": area = ob.Area; break;
                            default:                                             break;
			         }
               return area;
        }
        public string Extrair_texto                  (Entity enti)                    
        {
                      string texto = "-";
                      DBText txt   = enti as DBText;
                      texto        = txt.TextString;
                      return texto;
        }
        public string Extrair_Pontos_texto           (Entity enti)                    
        {
                      string coords = "-";
                      DBText entxt = enti as DBText;
                      Point3d p1   = entxt.Position;
                      coords = p1.X.ToString() + " " + p1.Y.ToString() + " " + p1.Z.ToString();
                      return coords;
        }
		public string Extrair_pontos_pontos          (Entity enti)                    
		{
			    string coords = ""; 
		        DBPoint pnt = enti as DBPoint;
			    Point3d p1  = pnt.Position; 
			    coords = p1.X.ToString() + " " + p1.Y.ToString() + " " + p1.Z.ToString();
			    return coords;
		}
		public string Extrair_pontos_linhas          (Entity enti)                    
		{
			    string coords = ""; 
		        Line linha = enti as Line;
			    Point3d p1 = linha.StartPoint; 
                Point3d p2 = linha.EndPoint; 
			    coords = p1.X.ToString() + " " + p1.Y.ToString() + " " + p1.Z.ToString() + " " + p2.X.ToString() + " " + p2.Y.ToString() + " " + p2.Z.ToString();
		  	    return coords;
		}
		public string Extrair_pontos_circle          (Entity enti)                    
		{
			       string coords  = ""; 
		           Circle circulo = enti as Circle;
			       Point3d p1 = circulo.Center;
			       coords = p1.X.ToString() + " " + p1.Y.ToString() + " " + p1.Z.ToString();
			       return coords;
		}
        public string Extrair_pontos_bloco           (Entity enti)                    
		{
			     string coords  = ""; 
		           BlockReference bloco = enti as BlockReference;
			       Point3d        p1    = bloco.Position;
			       coords         = p1.X.ToString() + " " + p1.Y.ToString() + " " + p1.Z.ToString();
			       return coords;
		}
		public string Extrair_pontos_poligonal2D     (Entity enti, double preci)                    
		{
			     string coords = ""; 
		         Polyline poly2d = enti as Polyline; 
			     double elevac = poly2d.Elevation;
                 List<Point2d> Lp = new List<Point2d> { };
                 for (int i = 0; i < poly2d.NumberOfVertices; i++)
                 {
                         Point2d p1 = poly2d.GetPoint2dAt(i);
                         if (Lp.Count == 0)
                             Lp.Add(p1);
                         else
                             { if (p1.GetDistanceTo(Lp.Last()) > preci)
                                  Lp.Add(p1);
                             }
                 }
                 
                 for (int i = 0; i < Lp.Count; i++)
                 {
                            Point2d p1 = Lp[i];
                            string vertice =  p1.X.ToString() + " " + p1.Y.ToString() + " " + elevac.ToString();
				            if (!coords.Contains(vertice))
				               {
				                   coords = coords + " " + vertice;
				               }
                 }
			  return coords;
		}
       
        public string Extrair_pontos_poligonal3D     (Entity enti, Transaction tr, double preci )   
		{
			            string coords = ""; 
   		                Polyline3d poly3d = enti as Polyline3d;
                        List<Point3d> Lp   = new List<Point3d> { };
                        List<double>  Lalt = new List<double> { };
                        foreach (ObjectId acObjIdVert in poly3d)
                        {
                              PolylineVertex3d p1 = tr.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d;
                              Lalt.Add(p1.Position.Z);
                        }
                        string alturamax = Lalt.Max().ToString();
            
                        foreach (ObjectId acObjIdVert in poly3d)
                        {
                                 PolylineVertex3d p1 = tr.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d;
                                 string vertice = p1.Position.X.ToString() + " " + p1.Position.Y.ToString() + " " + alturamax;
                                 if   (Lp.Count == 0)
                                      {
                                          coords = coords + " " + vertice;
                                          Lp.Add(p1.Position);
                                      }
                                 else {
                                          Point3d px = p1.Position;
                                          if (px.DistanceTo(Lp.Last()) > preci && !coords.Contains(vertice))
                                             {
				                                 coords = coords + " " + vertice;
                                                 Lp.Add(p1.Position);
                                             }
                                     }
                        }
			      return coords;
		}
     }

     public class Rio : Basicas
     {
        private List<string[]> barrial;  public List<string[]> Barrial { get { return barrial; } set { barrial = value; } }
        private List<string[]> tematic;  public List<string[]> Tematic { get { return tematic; } set { tematic = value; } }

        public void Cria_Texto   ( List<Point3d> Lpontos, string layer, string txt)  
		{
                                   Document acDoc = Application.DocumentManager.MdiActiveDocument;
                                   Database dbase = acDoc.Database;

                                   using (Transaction tr = dbase.TransactionManager.StartTransaction())
                                   {
                                          BlockTable       btl;
                                          BlockTableRecord btr;
                                          btl = tr.GetObject(dbase.BlockTableId, OpenMode.ForRead) as BlockTable;
                                          btr = tr.GetObject(btl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                          foreach (Point3d p in Lpontos)
                                          {
                                                 using  (DBText enttxt = new DBText())
                                                        {
                                                            btr.AppendEntity(enttxt);
                                                            tr.AddNewlyCreatedDBObject(enttxt, true);
                                                            enttxt.Layer      = layer;
                                                            enttxt.TextString = txt;
                                                            enttxt.Position   = p;
                                                            enttxt.Rotation   = 0.0;
                                                            enttxt.Height     = 1.5;
                                                            enttxt.Normal     = Vector3d.ZAxis;
                                                        }
                                          }
                                          tr.Commit();
                                   }
		}
        public void Cria_Ponto   ( List<Point3d> Lpontos, string layer )             
        {
                                   Document acDoc   = Application.DocumentManager.MdiActiveDocument;
                                   Database dbase = acDoc.Database;

                                   using (Transaction tr = dbase.TransactionManager.StartTransaction())
                                   {
                                          BlockTable       btl;
                                          BlockTableRecord btr;
                                          btl = tr.GetObject(dbase.BlockTableId, OpenMode.ForRead) as BlockTable;
                                          btr = tr.GetObject(btl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                          dbase.Pdmode = 34;
                                          dbase.Pdsize =  1;
                                          foreach (Point3d p in Lpontos)
                                          {
                                                  using (DBPoint entponto = new DBPoint(p))
                                                        {
                                                            btr.AppendEntity(entponto);
                                                            tr.AddNewlyCreatedDBObject(entponto, true);
                                                            entponto.Layer = layer;
                                                        }
                                          }
                                          tr.Commit();
                                   }
        }
        public void Cria_Linha   ( List<Point3d> Lpontos, string layer )             
        {
                                   Document acDoc = Application.DocumentManager.MdiActiveDocument;
                                   Database dbase = acDoc.Database;
                                   using (Transaction tr = dbase.TransactionManager.StartTransaction())
                                   {
                                          BlockTable       btl;
                                          BlockTableRecord btr;
                                          btl = tr.GetObject(dbase.BlockTableId, OpenMode.ForRead) as BlockTable;
                                          btr = tr.GetObject(btl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                                          Line lin = new Line(Lpontos[0], Lpontos[1]);
                                          tr.AddNewlyCreatedDBObject(lin, true);
                                          lin.Layer = layer;
                                          tr.Commit();
                                   }
        }
        public void Cria_Poli_3D ( List<Point3d> Lpontos, string layer , bool fecha) 
        {
                    Document acDoc = Application.DocumentManager.MdiActiveDocument;
                    Database dbase = acDoc.Database;
                    using (Transaction tr = dbase.TransactionManager.StartTransaction())
                    {
                           BlockTable       btl;
                           BlockTableRecord btr;
                           btl = tr.GetObject(dbase.BlockTableId, OpenMode.ForRead) as BlockTable;
                           btr = tr.GetObject(btl[BlockTableRecord.ModelSpace], OpenMode.ForWrite) as BlockTableRecord;
                           using (Polyline3d Poly3d = new Polyline3d())
                           {
                                    Poly3d.Closed = fecha;
                                    btr.AppendEntity(Poly3d);
                                    tr.AddNewlyCreatedDBObject(Poly3d, true);
                
                                    Point3dCollection acPts3dPoly = new Point3dCollection();
                                    for (int i = 0; i < Lpontos.Count; i++)   { acPts3dPoly.Add(Lpontos[i]); }
                                    foreach (Point3d acPt3d in acPts3dPoly)
                                    {
                                             using (PolylineVertex3d acPolVer3d = new PolylineVertex3d(acPt3d))
                                             {
                                                    Poly3d.AppendVertex(acPolVer3d);
                                                    tr.AddNewlyCreatedDBObject(acPolVer3d, true);
                                             }
                                    }
                                    Poly3d.Layer = layer;
                           }
                           tr.Commit();
                    }
        }
        public void Cria_Camada  ( string layer, short cor)                          
        {
               if (layer != "")
                  {
                     Document doc = Application.DocumentManager.MdiActiveDocument;
                     Database dba = doc.Database;
                     string   des = Rio_Dados_Tematica(layer)[2];
                     using (Transaction tra = dba.TransactionManager.StartTransaction())
                     {
                            LayerTable camadas = (LayerTable)tra.GetObject(dba.LayerTableId, OpenMode.ForRead);
                            if (!camadas.Has(layer))
                               {  
                                     camadas.UpgradeOpen();
                                     using (LayerTableRecord camareg = new LayerTableRecord())
                                         {
                                             camareg.Name        = layer;
                                             camareg.LineWeight  = LineWeight.LineWeight015;
                                             camareg.Color       = Color.FromColorIndex(ColorMethod.ByAci, cor);
                                             camadas.Add(camareg);
                                             tra.AddNewlyCreatedDBObject(camareg, true);
                                             camareg.Description = des;
                                         }
                              }
                              tra.Commit();
                     }
                  }
         }

// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
        [CommandMethod("RIOIN")]
        public void     Rio_IN()                                                                                        
        {
                            try
                                    {
                                      string       Pasta;
                                      List<string> Temario;
                                      List<string> Barrios;
                                      Rio_Selecion_Tema  (out Pasta, out Temario, out Barrios);
                                      if (Temario.Count > 0)
                                         {
                                                 Rio_Insertar_Temas(Pasta, Temario, Barrios);
                                         }
                                    }
                            catch   { }
                            finally { }
                            Mensaje( "Temas inseridos." ); 
        }
        public void     Rio_Selecion_Tema    ( out string Pasta, out List<string> Temario, out List<string> Barrios)    
        {
                          Pasta  = "D:\\JLMenegotto\\Academia\\04_Pesquisa\\MRJ\\MRJ_ORIG";
                          Form1 QDia = new Form1();
                          if (QDia.Temario.Count > 0)  { Temario = QDia.Temario; }  else  { Temario = new List<string> { "_" }; }
                          if (QDia.Barrios.Count > 0)  { Barrios = QDia.Barrios; }  else  { Barrios = new List<string> { "_" }; } 
                          Barrial = QDia.Rio_Barrial();
                          Tematic = QDia.Rio_GruTema();
        }
        public void     Rio_Insertar_Temas   ( string pasta,         List<string> TemasPro,    List<string> BarriosPro) 
        {
                                  for (int t = 0; t < TemasPro.Count; t++)
                                      {
                                             Cria_Camada(TemasPro[t], (short)(t + 10));  //Cria_os layers necessarios
                                      }
                                  for (int t = 0; t < TemasPro.Count; t++)
                                      {
                                             XmlDocument   xmlDoc       = new XmlDocument();
                                             string        Tema         = TemasPro[t];
                                             string        Arq_XML_Tema = pasta + "\\Rio_" + Tema + ".XML";
                                             xmlDoc.Load(Arq_XML_Tema);
                                             char[]        separa  = new char[] { ' ' };
                                             foreach (string barrio in BarriosPro)
                                             {
                                                     Rio_Procurar_Barrio(xmlDoc, Arq_XML_Tema, Tema, barrio);
                                             }
                                      }
        }
        public void     Rio_Procurar_Barrio  ( XmlDocument doc, string nomearq, string tema, string barr)               
        {
                            XmlNode     nodoraiz = null;
                            XmlNodeList L_nodos  = null;
                            XmlNodeList L_Elem   = null;
                            doc.Load(nomearq);
                            nodoraiz = doc.DocumentElement;
                            L_nodos = nodoraiz.SelectNodes("descendant::Bairro[@Bairro_Zona='" + barr + "']");
                            for (int i = 0; i < L_nodos.Count; i++)
                            {
                                    L_Elem = ((XmlElement)(L_nodos[i])).GetElementsByTagName("Elemento");
                                    for (int j = 0; j < L_Elem.Count; j++)
                                    {
                                         XmlElement elemento = (XmlElement)L_Elem[j];
                                         string   texto  =  "-";
                                         string   coord  = elemento.GetAttributeNode("Local").Value;
                                         string   objet  = elemento.GetAttributeNode("Objeto").Value;
                                         if (elemento.HasAttribute("Texto"))
                                            { texto  = elemento.GetAttributeNode("Texto").Value; }

                                         string[] coords = Rio_Separar_Coordena(coord.Substring(1));

                                         List<Point3d> Lpontos = new List<Point3d> { };
                                         for (int k = 0; k < (coords.Length / 3); k++)
                                         {  
                                               double x   = Convert.ToDouble(coords[3 * k + 0]);
                                               double y   = Convert.ToDouble(coords[3 * k + 1]);
                                               double z   = Convert.ToDouble(coords[3 * k + 2]);
                                               Point3d pt = new Point3d(x, y, z);
                                               Lpontos.Add(pt);
                                         }

                                         Rio_Desenhar_Objetos ( objet, tema, Lpontos, texto );
                                    }
                           } 
        }
        public void     Rio_Desenhar_Objetos ( string objeto, string tema, List<Point3d> Lpnts, string texto)           
        {
                     bool fecha = true;
                     switch (tema)
                     {
                                          case "338": fecha = false; break;
                                          case "339": fecha = false; break;
                                          case "500": fecha = false; break;
                                          case "502": fecha = false; break;
                                          case "503": fecha = false; break;
                                          case "504": fecha = false; break;
                                          case "505": fecha = false; break;
                                          case "506": fecha = false; break;
                                          case "507": fecha = false; break;
                                          case "508": fecha = false; break;
                                          case "562": fecha = false; break;
                                          case "572": fecha = false; break;
                                          case "650": fecha = false; break;
                                          case "651": fecha = false; break;
                                          case "652": fecha = false; break;
                                          case "653": fecha = false; break;
                                          case "654": fecha = false; break;
                                          case "655": fecha = false; break;
                                          case "656": fecha = false; break;
                                          case "657": fecha = false; break;
                                          case "658": fecha = false; break;
                                          case "671": fecha = false; break;
                                          case "672": fecha = false; break;
                                          case "696": fecha = false; break;
                                          case "714": fecha = false; break;
                                          case "720": fecha = false; break;
                                          case "721": fecha = false; break;
                                          case "722": fecha = false; break;
                                          case "723": fecha = false; break;
                                          case "780": fecha = false; break;
                                          case "781": fecha = false; break;
                                          case "782": fecha = false; break;
                                          case "783": fecha = false; break;
                                          case "784": fecha = false; break;
                                          case "785": fecha = false; break;
                                          case "786": fecha = false; break;
                                          default: break;
                     }
                     switch (objeto)
                     {
                                          case "Ponto":   Cria_Ponto(Lpnts, tema);          break;
                                          case "Bloco":   Cria_Ponto(Lpnts, tema);          break;
                                          case "Circulo": Cria_Ponto(Lpnts, tema);          break;
                                          case "Linha":   Cria_Linha(Lpnts, tema);          break;
                                          case "Poli_2D": Cria_Poli_3D(Lpnts, tema, fecha); break;
                                          case "Poli_3D": Cria_Poli_3D(Lpnts, tema, fecha); break;
                                          case "Texto":   Cria_Texto  (Lpnts, tema, texto); break;
                                          default: break;
                     }
        }
        public string[] Rio_Separar_Coordena ( string coord)                                                            
        {
                            char[] separa = new char[] { ' ' };
                            return coord.Split(separa);
        }

// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
        [CommandMethod("RIOOUT")]
        public void RIO_OUT()   
        {
                            XmlWriterSettings opc     = new XmlWriterSettings();
                            opc.Indent                = true;
                            opc.CheckCharacters       = false;
                            List<XmlWriter> Lista_xml = new List<XmlWriter> { };
                            try     { string        Pasta;
                                      List<string>  Temario;
                                      List<string>  Barrios;
                                      Rio_Selecion_Tema(out Pasta, out Temario, out Barrios);
                                      if (Temario.Count > 0)
                                      {
                                           for (int t = 0; t < Temario.Count; t++)
                                           {
                                                string Tema         = Temario[t];
                                                string Arq_XML_Tema = Pasta + "\\Rio_" + Tema + ".XML";
                                                XmlWriter  XML      = XmlWriter.Create(Arq_XML_Tema, opc);
                                                Rio_Elem_Cidade( XML , Tema );
                                                Lista_xml.Add(   XML  );      //Lista com os Temas XML iniciados
                                           }
                                           Rio_Explorando_DWGs( Pasta, Temario, Barrios, Lista_xml );
                                      }
                                      Rio_Fecha_Cidade   ( Lista_xml ); //Fecha a cidade para cada tema
                                    }
                            catch   { }
                            finally { }
                    Mensaje("Temas Extraidos e Inseridos.");
        }

        [CommandMethod("RIOOUT_IN")]
        public void RIOOUTIN()  
        {
                            XmlWriterSettings opc     = new XmlWriterSettings();
                            opc.Indent                = true;
                            opc.CheckCharacters       = false;
                            List<XmlWriter> Lista_xml = new List<XmlWriter> { };
                            try     { string        Pasta;
                                      List<string>  Temario;
                                      List<string>  Barrios;
                                      Rio_Selecion_Tema(out Pasta, out Temario, out Barrios);
                                      if (Temario.Count > 0)
                                      {
                                           for (int t = 0; t < Temario.Count; t++)
                                           {
                                                string Tema         = Temario[t];
                                                string Arq_XML_Tema = Pasta + "\\Rio_" + Tema + ".XML";
                                                XmlWriter  XML      = XmlWriter.Create(Arq_XML_Tema, opc);
                                                Rio_Elem_Cidade( XML , Tema );
                                                Lista_xml.Add(   XML  );      //Lista com os Temas XML iniciados
                                           }
                                           Rio_Explorando_DWGs( Pasta, Temario, Barrios, Lista_xml );
                                      }
                                      Rio_Fecha_Cidade   ( Lista_xml ); //Fecha a cidade para cada tema
                                      if (Temario.Count > 0)
                                      {
                                          Rio_Insertar_Temas(Pasta, Temario, Barrios);
                                      }
                                    }
                            catch   { }
                            finally { }
                    Mensaje("Temas Extraidos e Inseridos.");


        }

        public void         Rio_Explorando_DWGs ( string pasta, List<string> TemasPro, List<string> BarriosPro, List<XmlWriter> LXML) 
        {
                            List<string> Arquivos = Rio_Montar_Barrios( BarriosPro );
                            try     {
                                         for (int a = 0; a < Arquivos.Count; a++)
                                         {
                                             string dwg = Arquivos[a];                                                   
                                             Rio_Processa_Barrio( pasta, TemasPro , dwg , LXML);
                                             Rio_Fecha_Bairros  ( LXML ); //fecha o bairro para cada arquivo
                                         }
                                    }
                            catch   { }
                            finally { }

        }
        public List<string> Rio_Montar_Barrios  ( List<string> BarriosPro)                                                            
        {
                              List<string> procesarArquivosDWG = new List<string>{};
                              foreach (string b in BarriosPro)
                              {
                                      IEnumerable<string> bb = Barrial.Where(strings => strings.Contains(b)).Select(strings => strings[2]);
                                      foreach (string s in bb) { procesarArquivosDWG.Add( s ); }
                              }
                              return procesarArquivosDWG;
        }
        public void         Rio_Processa_Barrio ( string pasta, List<string> Temas   , string ArqDWG,           List<XmlWriter> LXML) 
        {
                     string Arqudwg = pasta + "\\" + ArqDWG;
                     for (int t = 0; t < Temas.Count; t++) { Rio_Elem_Bairro (LXML[t] , ArqDWG); }

                     Database dbside = new Database(false, true);
                     using (dbside)
                     {
                            dbside.ReadDwgFile(Arqudwg, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                            Transaction tr = dbside.TransactionManager.StartOpenCloseTransaction();
                            using (tr)
                            {
                                BlockTable       bt   = (BlockTable)tr.GetObject(dbside.BlockTableId, OpenMode.ForRead);
                                BlockTableRecord btr  = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                                foreach (ObjectId objId in btr)
                                {
                                    Entity entid = (Entity)tr.GetObject(objId, OpenMode.ForRead);
                                    string tpobj = Tipo_de_Objeto(entid);
                                    string texto = Texto_do_Objeto(entid);
                                    string layer = entid.Layer.Substring(0, 3);
                                    string coord = Extrair_Coordenadas(entid, tr);
                                    for (int t = 0; t < Temas.Count; t++)
                                    {
                                             XmlWriter XML   =  LXML[t];  //Arquivo XML do Tema
                                             string   bair   = ArqDWG.Substring(0, ArqDWG.Length - 4);
                                             string   Tema   = Temas[t];
                                             if (tpobj != "Desconocido" && layer.Contains(Tema))
                                                {
                                                      Rio_Elem_Elemen ( XML , tpobj , coord , texto, Tema , bair); 
                                                }
                                    }
                               } 
                               tr.Commit();
                            }
                            dbside.CloseInput(true);
                            dbside.Dispose();
                      } 
        } 
// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
// ---------------------- FUNÇÕES DE CRIAÇÃO E MODIFICAÇÃO DO XML -------------------------------------------
// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
        public void         Rio_Elem_Cidade   ( XmlWriter XML,  string Tema)              
        {
                      XML.WriteStartElement("Cidade");
                      XML.WriteAttributeString("Nome_Cidade", "Rio de Janeiro");
                      XML.WriteAttributeString("Tema",      Tema);
                      XML.WriteAttributeString("Descrição", Rio_Dados_Tematica(Tema)[2]);
        } 
        public void         Rio_Elem_Bairro   ( XmlWriter XML,  string ADWG)              
        {
                            List<string> dados = Rio_Dados_Bairro ( ADWG );
                            XML.WriteStartElement("Bairro");
                                 XML.WriteAttributeString("Bairro_AP",      dados[0]);
                                 XML.WriteAttributeString("Bairro_RA",      dados[1]);
                                 XML.WriteAttributeString("Bairro_Arquivo", dados[2]);
                                 XML.WriteAttributeString("Bairro_Nome",    dados[3]);
                                 XML.WriteAttributeString("Bairro_Zona",    dados[4]);
                                 XML.WriteAttributeString("Bairro_Codigo",  dados[5]);
        }
        public void         Rio_Elem_Elemen   ( XmlWriter XML,  string obtp, string coor, string txt, string tema, string bairr) 
        {
                            XML.WriteStartElement("Elemento");
                                                        XML.WriteAttributeString("Tema"   , tema);
                                                        XML.WriteAttributeString("Bairro" , bairr);
                                                        XML.WriteAttributeString("Objeto" , obtp);
                                                        XML.WriteAttributeString("Local"  , coor);
                                   if (obtp == "Texto") XML.WriteAttributeString("Texto"  , txt);
                            XML.WriteEndElement();
        }
        public void         Rio_Fecha_Bairros ( List<XmlWriter> Lista_xml )               
        { 
                            for (int t = 0; t < Lista_xml.Count; t++)
                                {
                                       XmlWriter XML = Lista_xml[t];
                                       XML.WriteEndElement(); //fecha bairros  
                                }
        }
        public void         Rio_Fecha_Cidade  ( List<XmlWriter> Lista_xml )               
        { 
                            for (int t = 0; t < Lista_xml.Count; t++)
                                {
                                       XmlWriter XML = Lista_xml[t];
                                       XML.WriteEndElement(); //fecha cidade 
                                       XML.Flush();
                                       XML.Close();
                                }
       }
        public List<string> Rio_Dados_Bairro  ( string barrio )                           
        {
                            List<string> dados = Barrial.Where(strings => strings.Contains(barrio)).Single().ToList();
                            return dados;
        }
        public List<string> Rio_Dados_Tematica( string tema)                              
        {
                           List<string> dados = Tematic.Where(strings => strings.Contains(tema)).Single().ToList();
                           return dados;
        }
        public List<string> Rio_Lst_Arq_DWG   ( string pasta )                            
        {
                            DirectoryInfo   contpasta  = new DirectoryInfo  ( pasta );
                            FileInfo[] arquivos = contpasta.GetFiles("*.dwg");
                            List<string> L_arquivos = new List<string> { };
                            foreach (FileInfo arquivo in arquivos)
                            {
                                    L_arquivos.Add(arquivo.Name.ToLower());
                            }
                            return L_arquivos;
       }
     



        [CommandMethod("TESTE")]
        public void TESTE()                                      
        {

        }
        [CommandMethod("seleccion")]
        public void Seleccion(Editor edi, Point3dCollection Lpts)
        {
               Editor edoc = Application.DocumentManager.MdiActiveDocument.Editor;
               PromptSelectionResult resultado;
               resultado = edi.SelectCrossingPolygon(Lpts);
               if (resultado.Status == PromptStatus.OK)
                  {
                             SelectionSet sset = resultado.Value;
                             Application.ShowAlertDialog("Objetos selecionados: " + sset.Count.ToString());
                  }
                else
                  {
                             Application.ShowAlertDialog("Number of objects selected: 0");
                  }
        }
        [CommandMethod("procesar")]
        public void Rio_Processa_DWGs ( )                        
        {

                     Point3dCollection Lp = new Point3dCollection { new Point3d(0, 0, 0), new Point3d(20, 0, 0), new Point3d(20, 20, 0), new Point3d(0, 20, 0) };
                     string        Pasta = "D:\\JLMenegotto\\Academia\\04_Pesquisa\\MRJ\\";
                     List<string> Arquivos = new List<string> { "Teste1.dwg", "Teste2.dwg", "Teste3.dwg", "Teste4.dwg" };
                     Database dbside = new Database(false, true);
                     foreach (string arq in Arquivos)
                     {
                                string abrir = Pasta + arq;
                                using (dbside)
                                {
                                       dbside.ReadDwgFile(abrir, FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                                       Document doc = Application.DocumentManager.CurrentDocument;
                                       Editor   edi = doc.Editor;
                                       Seleccion(edi, Lp);
                                }
                                dbside.CloseInput(true);
                     }
                     dbside.Dispose();
        }
     }
}
