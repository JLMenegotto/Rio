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
using G = CoordinateSharp;

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

        private string _drive;
        public string   Drive   { get { return _drive; } set { _drive = value; } }

        public Document    doc  { get { return Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument; }                         set { _doc = value; } }
		public Editor      edi  { get { return Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor; }                  set { _edi = value; } }
		public Database    dba  { get { return Autodesk.AutoCAD.ApplicationServices.Core.Application.DocumentManager.MdiActiveDocument.Editor.Document.Database;} set { _db1 = value; } }
		public Database    dbw  { get { return HostApplicationServices.WorkingDatabase; }                                                                         set { _db2 = value; } }
		public Transaction tran { get { return _tran; }                                                                                                           set { _tran = value; } }

		public double       Str_Doble (string s) { return System.Convert.ToDouble(s);}
		public long         Str_Int64 (string s) { return System.Convert.ToInt64(s); }
		public int          Str_Int32 (string s) { return System.Convert.ToInt32(s); }

        public void         Mensaje    (string m)                          
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
		public DBObject Cria_Esfera  ( Transaction tr, Point3d centro, double raio, string nomelayer ) 
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
		public DBObject Cria_Circulo ( Transaction tr, Point3d centro, double raio )                   
		{
			Circle cir = new Circle();
			cir.Center = centro;
			cir.Radius = raio;
			cir.Normal = Vector3d.ZAxis;
			NovaEntidad(tr, cir);
			DBObject obj = tr.GetObject(cir.ObjectId, OpenMode.ForRead);
			return obj;
		}
        public DBObject Cria_Circulo ( Transaction tr, Point3d centro, double raio, string nomelayer ) 
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
		public DBObject Cria_Texto   ( Transaction tr, Point3d pto, string txt )                       
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
		public DBObject Cria_Hachura ( Transaction tr, DBObject obj, double elev )                     
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
		public Entity   NovaEntidad  ( Transaction tr, Entity  ent )                                   
		{
                        BlockTableRecord btr = (BlockTableRecord)tr.GetObject(dba.CurrentSpaceId, OpenMode.ForWrite);
                        ObjectId obj = btr.AppendEntity(ent);
                        tr.AddNewlyCreatedDBObject(ent, true);
                        return ent;
		}
        public void     Faz_Offset   ( Transaction tr,  Polyline pl   , BlockTableRecord btr)          
        {
                        foreach (Entity ent in pl.GetOffsetCurves(0.025))
                        {
                                 btr.AppendEntity( ent );
                                 tr.AddNewlyCreatedDBObject( ent , true );
                        }
        }

		public ObjectId[]   Prepara_Zonas     ( )                                       
		{
			ObjectId[] PoligonaisZonas = Captura_Zonas("LWPOLYLINE", true);
			return PoligonaisZonas;
		}
		public string       Ponto_Dentro      ( Polyline poli, Point3d p0 )             
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
		public double       AreaTriangulo     ( Point3d  p1,   Point3d p2, Point3d p3 ) 
		{
			double a = p1.DistanceTo(p2);
			double b = p2.DistanceTo(p3);
			double c = p3.DistanceTo(p1);
			double S = (a + b + c) / 2.0;
			double area = Math.Sqrt(S * (S - a) * (S - b) * (S - c));
			return area;
		}
		public Point3d      Seleccion_Ponto   ( string msg )                            
		{
			PromptPointOptions pop = new PromptPointOptions("\n" + msg);
			PromptPointResult pnt = edi.GetPoint(pop);
			if (pnt.Status == PromptStatus.OK)
				return pnt.Value;
			else
				return new Point3d();
		}
             
		public int          Ingressar_Inteiro ( string msg, int    defaul)              
		{
			                PromptIntegerResult numint;
		                    PromptIntegerOptions inte = new PromptIntegerOptions("\n" + msg);
                            inte.DefaultValue         = defaul;
                            inte.UseDefaultValue      = true;
                            numint                    = edi.GetInteger(inte);
                     return numint.Value;
		}
		public double       Ingressar_Real    ( string msg, double defaul )             
		{
			                PromptDoubleResult numreal;
			                PromptDoubleOptions real = new PromptDoubleOptions("\n" + msg);
			                real.DefaultValue        = defaul;
			                real.UseDefaultValue     = true;
			                numreal                  = edi.GetDouble(real);
			         return numreal.Value;
		}
		public string       Ingressar_Text    ( string msg, string defaul )             
		{
			                PromptResult texto;
			                PromptStringOptions textoin = new PromptStringOptions("\n" + msg);
			                textoin.DefaultValue        = defaul;
			                textoin.UseDefaultValue     = true;
			                texto                       = edi.GetString(textoin);
                            if (texto.Status != PromptStatus.OK)
                                 return defaul;
                             else
                                 return texto.StringResult;
		}

		public bool         ParImp(int num)          { return num % 2   == 0; }             // Verifica se número é par ou impar       
		public bool         Modulo(int num, int mod) { return num % mod == 0; }             // Verifica indice modular do número       
		public int          Congru(int num, int mod)                                    
		{
                            int partes  = (num / mod);
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
		
		public string   Extrair_Pontos     ( Entity enti, Transaction tr)                            
		{
                      string t = enti.GetType().ToString();
                      string p = "-";
                      string s = " ";
                      switch (t)
                      {
                          case "Autodesk.AutoCAD.DatabaseServices.DBPoint":        p = Extrair_xyz_pontos ( enti , s          ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Line":           p = Extrair_xyz_linhas ( enti , s          ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Circle":         p = Extrair_xyz_circle ( enti , s          ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Polyline":       p = Extrair_xyz_poly2D ( enti , s , 0.1    ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Polyline2d":     p = Extrair_xyz_poly2D ( enti , s , 0.1    ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Polyline3d":     p = Extrair_xyz_poly3D ( enti , s , 0.1, tr); break;
                          case "Autodesk.AutoCAD.DatabaseServices.BlockReference": p = Extrair_xyz_bloco  ( enti , s          ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.DBText":         p = Extrair_xyz_texto  ( enti , s          ); break;
                          default:                                                 p = "-";                                      break;
                      }
               return p;
		}
        public string   Extrair_Pontos_KML ( Entity enti, Transaction tr)                            
        {
                      string t = enti.GetType().ToString();
                      string p = "-";
                      switch (t)
                      {
                          case "Autodesk.AutoCAD.DatabaseServices.DBPoint":        p = Extrair_UTM_pontos (enti           ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Line":           p = Extrair_UTM_linhas (enti           ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Circle":         p = Extrair_UTM_circle (enti           ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Polyline":       p = Extrair_UTM_poly2D (enti , 0.1     ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Polyline2d":     p = Extrair_UTM_poly2D (enti , 0.1     ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.Polyline3d":     p = Extrair_UTM_poly3D (enti , 0.1 , tr); break;
                          case "Autodesk.AutoCAD.DatabaseServices.BlockReference": p = Extrair_UTM_bloco  (enti           ); break;
                          case "Autodesk.AutoCAD.DatabaseServices.DBText":         p = Extrair_UTM_texto  (enti           ); break;
                          default:                                                 p = "-";                                  break;
                      }
               return p;
        }
        public string   Tipo_de_Objeto     ( Entity enti)                                            
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
        public string   Texto_do_Objeto    ( Entity enti)                                            
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
        public double   Area_objeto        ( Entity enti)                                            
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

        public string   Extrair_texto      ( Entity enti)                                            
        {
                      string texto = "-";
                      DBText txt   = enti as DBText;
                      texto        = txt.TextString;
                      return texto;
        }
        public string   Extrair_xyz_texto  ( Entity enti, string sep)                                
        {
                      string coords = "-";
                      DBText entxt  = enti as DBText;
                      Point3d p1    = entxt.Position;
                      coords        = p1.X.ToString() + sep + p1.Y.ToString() + sep + p1.Z.ToString();
                      return coords;
        }
		public string   Extrair_xyz_pontos ( Entity enti, string sep)                                
		{
			           string  coords = ""; 
		               DBPoint pnt    = enti as DBPoint;
			           Point3d p1     = pnt.Position; 
			           coords         = p1.X.ToString() + sep + p1.Y.ToString() + sep + p1.Z.ToString();
			    return coords;
		}
		public string   Extrair_xyz_linhas ( Entity enti, string sep)                                
		{
			          string  coords = ""; 
		              Line    linha  = enti as Line;
			          Point3d p1     = linha.StartPoint; 
                      Point3d p2     = linha.EndPoint; 
			          coords         = p1.X.ToString() + sep + p1.Y.ToString() + sep + p1.Z.ToString() + " " + p2.X.ToString() + sep + p2.Y.ToString() + sep + p2.Z.ToString();
		  	          return coords;
		}
		public string   Extrair_xyz_circle ( Entity enti, string sep)                                
		{
			          string  coords  = ""; 
		              Circle  circulo = enti as Circle;
			          Point3d p1      = circulo.Center;
			          coords          = p1.X.ToString() + sep + p1.Y.ToString() + sep + p1.Z.ToString();
			          return coords;
		}
        public string   Extrair_xyz_bloco  ( Entity enti, string sep)                                
		{
			          string         coords = ""; 
		              BlockReference bloco  = enti as BlockReference;
			          Point3d        p1     = bloco.Position;
			          coords                = p1.X.ToString() + sep + p1.Y.ToString() + sep + p1.Z.ToString();
			          return coords;
		}
		public string   Extrair_xyz_poly2D ( Entity enti, string sep, double preci)                  
		{
			     string        coords  = ""; 
		         Polyline      poly2d  = enti as Polyline; 
			     double        elevac  = poly2d.Elevation;
                 List<Point2d> Lp      = new List<Point2d> { };
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
                            string vertice =  p1.X.ToString() + sep + p1.Y.ToString() + sep + elevac.ToString();
				            if (!coords.Contains(vertice))
				               {
				                   coords = coords + " " + vertice;
				               }
                 }
			  return coords;
		}
        public string   Extrair_xyz_poly3D ( Entity enti, string sep, double preci , Transaction tr) 
		{
			            string        coords = ""; 
   		                Polyline3d    poly3d = enti as Polyline3d;
                        List<Point3d> Lp     = new List<Point3d> { };
                        List<double>  Lalt   = new List<double>  { };
                        foreach (ObjectId acObjIdVert in poly3d)
                        {
                              PolylineVertex3d p1 = tr.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d;
                              Lalt.Add(p1.Position.Z);
                        }
                        string alturamax = Lalt.Max().ToString();
            
                        foreach (ObjectId acObjIdVert in poly3d)
                        {
                                 PolylineVertex3d p1 = tr.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d;
                                 string vertice = p1.Position.X.ToString() + sep + p1.Position.Y.ToString() + sep + alturamax;
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

        public Vector2d VetorCorrecao2d = new Vector2d(-43.85591336, -45.81013948);
        public Vector3d VetorCorrecao3d = new Vector3d(-43.85591336, -45.81013948, 0);
        public string   ElevacaoPonto = "";
        public string   ElevacaoTotal = "";

        public string   Traduz_GeoPonto    ( Point3d p1 )                                            
        {
                        double altura = p1.Z;
                        if (altura == 0)
                            altura = 6.1234;
                        
                        Point3d pc = p1.Add( VetorCorrecao3d ); //desloca para corrigir posição
                        G.UniversalTransverseMercator utm1 = new G.UniversalTransverseMercator("23K", pc.X, pc.Y);
                        double[] lat1 = G.UniversalTransverseMercator.ConvertUTMtoSignedDegree(utm1);
                        string   coor = lat1[1].ToString() + "," + lat1[0].ToString() + "," + altura.ToString();
                 return coor;
        }
        public string   Traduz_GeoPonto    ( Point2d p1 ,           double altura)                   
        {
                         if (altura == 0)
                            altura = 6.1234; 

                        Point2d pc = p1.Add( VetorCorrecao2d ); //desloca para corrigir posição
                        G.UniversalTransverseMercator utm1 = new G.UniversalTransverseMercator("23K", pc.X, pc.Y);
                        double[] lat1 = G.UniversalTransverseMercator.ConvertUTMtoSignedDegree(utm1);
                        string coor = lat1[1].ToString() + "," + lat1[0].ToString() + "," + altura.ToString();
                 return coor;
        }
        public string   Traduz_GeoPonto    ( PolylineVertex3d p1  , double altura)                   
        {
                        if (altura == 0)
                            altura = 6.1234;

                        Point3d pc = new Point3d(p1.Position.X, p1.Position.Y, p1.Position.Z).Add( VetorCorrecao3d ); 
                        G.UniversalTransverseMercator utm1 = new G.UniversalTransverseMercator("23K", pc.X, pc.Y);
                        double[] lat1 = G.UniversalTransverseMercator.ConvertUTMtoSignedDegree(utm1);
                        string   coor = lat1[1].ToString() + "," + lat1[0].ToString() + "," + altura.ToString();
                 return coor;
        }
        public string   Extrair_UTM_texto  ( Entity  enti )                                          
        {
                      DBText entxt  = enti as DBText;
                      Point3d p1    = entxt.Position;
                      string coords = Traduz_GeoPonto(p1);
               return coords;
        }
        public string   Extrair_UTM_pontos ( Entity  enti )                                          
        {
                      DBPoint pnt   = enti as DBPoint;
                      Point3d p1    = pnt.Position;
                      string coords = Traduz_GeoPonto( p1 );
               return coords;
        }
        public string   Extrair_UTM_linhas ( Entity  enti )                                          
        {
                      Line linha   = enti as Line;
                      Point3d p1   = linha.StartPoint;
                      Point3d p2   = linha.EndPoint;
                      string  coo1 = Traduz_GeoPonto( p1 );
                      string  coo2 = Traduz_GeoPonto( p2 );
               return coo1 + " " + coo2;
        }
        public string   Extrair_UTM_circle ( Entity  enti )                                          
        {
                      Circle circulo = enti as Circle;
                      Point3d p1     = circulo.Center;
                      string coords  = Traduz_GeoPonto( p1 );
               return coords;
        }
        public string   Extrair_UTM_bloco  ( Entity  enti )                                          
        {
                      BlockReference bloco  = enti as BlockReference;
                      Point3d        p1     = bloco.Position;
                      string         coords = Traduz_GeoPonto(p1);
               return coords;
        }
        public string   Extrair_UTM_poly2D ( Entity enti, double preci)                              
        {
                      string        coords = "";
                      Polyline      poly2d = enti as Polyline;
                      double        elevac = poly2d.Elevation;
            
                      List<Point2d> Lp     = new List<Point2d> { };
                      for (int i = 0; i < poly2d.NumberOfVertices; i++)
                      {
                               Point2d p1 = poly2d.GetPoint2dAt(i);
                               if (Lp.Count == 0)
                                   Lp.Add(p1);
                               else
                               {
                                   if (p1.GetDistanceTo(Lp.Last()) > preci)
                                   Lp.Add ( p1 );
                               }
                      }

                      Lp.Add(Lp.First());

                      for (int i = 0; i < Lp.Count; i++)
                      {
                               Point2d p1  = Lp[i];
                               string vrtz = Traduz_GeoPonto( p1 , elevac );
                               if (!coords.Contains( vrtz ))
                               {
                                    coords = coords + " " + vrtz;
                               }
                      }
               return coords;
        }
        public string   Extrair_UTM_poly3D ( Entity enti, double preci, Transaction tr)              
        {      
                      string                 coords = "";
                      Polyline3d             poly3d = enti as Polyline3d;
                      List<PolylineVertex3d> Lvrtcs = new List<PolylineVertex3d> { };
                      List<string>           Lvzgeo = new List<string>  { };
                      List<double>           Lalt   = new List<double>  { };

                      foreach (ObjectId acObjIdVert in poly3d)
                      {
                              PolylineVertex3d vrtc1 = tr.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d;
                              Lalt.Add( vrtc1.Position.Z );
                      }

                      double  alturamin = Lalt.Min();
                      double  alturamax = Lalt.Max();

                      foreach (ObjectId acObjIdVert in poly3d)
                      {
                               PolylineVertex3d vrtc1 = tr.GetObject(acObjIdVert, OpenMode.ForRead) as PolylineVertex3d;
                               string           vrtz  = Traduz_GeoPonto( vrtc1 , alturamax );
                               Lvzgeo.Add(vrtz);

                               if (Lvrtcs.Count == 0)
                               {
                                       coords = coords + " " + vrtz;
                                       Lvrtcs.Add( vrtc1 );
                               }
                               else
                               {     
                                    if (vrtc1.Position.DistanceTo (Lvrtcs.Last().Position) > preci && !coords.Contains(vrtz))
                                    {
                                       coords = coords + " " + vrtz;
                                       Lvrtcs.Add( vrtc1 );
                                    }
                               }
                      }
                      coords = coords + " " + Lvzgeo.First(); //adiciona o primeiro para fechar o ring
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
                     string   des = Rio_Dados_Tematica(layer)[4];
                     using (Transaction tra = dba.TransactionManager.StartTransaction())
                     {
                           LayerTable camadas = (LayerTable)tra.GetObject(dba.LayerTableId, OpenMode.ForRead);
                           if (!camadas.Has(layer))
                           {  
                                camadas.UpgradeOpen();
                                using (LayerTableRecord camareg = new LayerTableRecord())
                                {
                                       camareg.Name       = layer;
                                       camareg.LineWeight = LineWeight.LineWeight015;
                                       camareg.Color      = Color.FromColorIndex(ColorMethod.ByAci, cor);
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
                          string Drive = "C";
                          Pasta      = Drive + ":\\JLMenegotto\\Academia\\04_Pesquisa\\MRJ\\MRJ_ORIG";
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
        public void     Rio_Desenhar_Objetos ( string objeto  , string tema, List<Point3d> Lpnts, string texto)         
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
        public List<string> Rio_Montar_Barrios ( List<string> Barrios) 
        {
                              List<string> procesarArquivosDWG = new List<string>{};
                              foreach (string b in Barrios)
                              {
                                      IEnumerable<string> bb = Barrial.Where(strings => strings.Contains(b)).Select(strings => strings[2]);
                                      foreach (string s in bb) { procesarArquivosDWG.Add( s ); }
                              }
                              return procesarArquivosDWG;
        }
        public List<string> Rio_Dados_Bairro   ( string barrio )       
        {
                            List<string> dados = Barrial.Where(strings => strings.Contains(barrio)).Single().ToList();
                            return dados;
        }
        public List<string> Rio_Dados_Tematica ( string tema   )       
        {
                           List<string> dados = Tematic.Where(strings => strings.Contains(tema)).Single().ToList();
                           return dados;
        }
        public List<string> Rio_Lst_Arq_DWG    ( string pasta  )       
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

        [CommandMethod("RIOCOORD")]
        public void RIOCOORD (  )   
        {

            Point3d p1 = new Point3d(683843.4700 , 7457149.1300 , 11.8000); //google_earth UTM = 683800.27 m E  7457104.07 m S
            Point3d p2 = new Point3d(684015.5400 , 7457137.6200 , 8.1000 );
            Point3d p3 = new Point3d(684025.4000 , 7457037.1100 , 14.8000);
            Point3d p4 = new Point3d(683827.9700 , 7457050.3500 , 14.9000);

            G.UniversalTransverseMercator utm1   = new G.UniversalTransverseMercator("23K", p1.X, p1.Y);
            G.Coordinate                  lat1_a = G.UniversalTransverseMercator.ConvertUTMtoLatLong(utm1);
            double[]                      lat1_b = G.UniversalTransverseMercator.ConvertUTMtoSignedDegree(utm1);

            Mensaje (  "683843.4700 , 7457149.1300 , 11.8000"   + "\n" +
                       "LatLong      = " + lat1_a.ToString()    + "\n" +
                       "SignedDegree = " + lat1_b[0].ToString() + "  " + lat1_b[1].ToString()
                    );
        }

        [CommandMethod("RIOOUT_IN")]
        public void RIOOUTIN (  )   
        {
                            XmlWriterSettings opc     = new XmlWriterSettings();
                            opc.Indent                = true;
                            opc.CheckCharacters       = false;
                            List<XmlWriter> L_xml = new List<XmlWriter> { };
                            try     {
                                       string        Pasta;
                                       List<string>  Temario;
                                       List<string>  Barrios;
                                       Rio_Selecion_Tema (out Pasta, out Temario, out Barrios);
                                       if (Temario.Count > 0)
                                       {
                                           for (int t = 0; t < Temario.Count; t++)
                                           {
                                                string Tema     = Temario[t];
                                                string Arq_Tema = Pasta + "\\XML_Rio_" + Tema + ".XML";
                                                XmlWriter  XML  = XmlWriter.Create(Arq_Tema, opc);
                                                Rio_Elem_Cidade( XML , Tema );
                                                L_xml.Add      ( XML        ); //Lista com Temas XML iniciados
                                           }
                                           Rio_Explorar_DWG( Pasta, Temario, L_xml , Barrios);
                                       }
                                       Rio_Fecha_Cidade   ( L_xml ); //Fecha a cidade para cada tema
                                       if (Temario.Count > 0)
                                       {
                                           Rio_Insertar_Temas(Pasta, Temario, Barrios);
                                       }
                                    }
                            catch   { }
                            finally { }
                    Mensaje("Temas Extraidos e Inseridos.");
        }

// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------
// ---------------------- FUNÇÕES DE CRIAÇÃO E MODIFICAÇÃO DO XML -------------------------------------------
// ----------------------------------------------------------------------------------------------------------
// ----------------------------------------------------------------------------------------------------------   
        [CommandMethod("RIOXML")]
        public void  RIO_XML()                                                                                             
        {
                            XmlWriterSettings opc = new XmlWriterSettings();
                            opc.Indent            = true;
                            opc.CheckCharacters   = false;
                            List<XmlWriter> L_xml = new List<XmlWriter> { };
                            try     { string        Pasta;
                                      List<string>  Temario;
                                      List<string>  Barrios;
                                      Rio_Selecion_Tema(out Pasta, out Temario, out Barrios);
                                      if (Temario.Count > 0)
                                      {
                                            for (int t = 0; t < Temario.Count; t++)
                                            {
                                                  string    Tema         = Temario[t];
                                                  string    Arq_XML_Tema = Pasta + "\\XML_Rio_" + Tema + ".XML";
                                                  XmlWriter XML          = XmlWriter.Create(Arq_XML_Tema, opc);
                                                  Rio_Elem_Cidade( XML , Tema );
                                                  L_xml.Add      ( XML        );   //Lista com os Temas XML iniciados
                                            }
                                            Rio_Explorar_DWG ( Pasta, Temario , L_xml , Barrios);  //faz so kml
                                      }
                                      Rio_Fecha_Cidade ( L_xml ); //Fecha a cidade para cada tema
                                    }
                            catch   { }
                            finally { }
                    Mensaje("Temas extraidos em arquivos XML.");
        }
        public void  Rio_Processa_Bar  ( string pasta  , List<string> Temas , List<XmlWriter> LXML ,      string ArqDWG )  
        {
                     string Arqudwg = pasta + "\\" + ArqDWG;
                     for (int t = 0; t < Temas.Count; t++)
                     {
                          Rio_Elem_Bairro ( LXML[t] , ArqDWG );
                     }

                     Database dbside = new Database(false, true);
                     using (dbside)
                     {
                            dbside.ReadDwgFile ( Arqudwg , FileOpenMode.OpenForReadAndWriteNoShare , true , "" );

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
                                          string coord = Extrair_Pontos(entid, tr);
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
        public void  Rio_Explorar_DWG  ( string pasta  , List<string> Temas , List<XmlWriter> LXML , List<string> Barrios) 
        {
                            List<string> Arquivos = Rio_Montar_Barrios( Barrios );
                            try
                            {
                                         for (int a = 0; a < Arquivos.Count; a++)
                                         {
                                                string dwg = Arquivos[a];
                                                Rio_Processa_Bar  ( pasta, Temas , LXML, dwg);
                                                Rio_Fecha_Bairros ( LXML ); //fecha o bairro para cada arquivo
                                         }
                            }
                            catch   { }
                            finally { }

        }     
        public void  Rio_Elem_Cidade   ( XmlWriter XML , string Tema )                                                     
        {
                            XML.WriteStartElement("Cidade");
                            XML.WriteAttributeString("Nome_Cidade", "Rio de Janeiro");
                            XML.WriteAttributeString("Tema",      Tema);
                            XML.WriteAttributeString("Descrição", Rio_Dados_Tematica(Tema)[4]);
        }
        public void  Rio_Elem_Bairro   ( XmlWriter XML , string ADWG )                                                     
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
        public void  Rio_Elem_Elemen   ( XmlWriter XML , string ob, string xy, string tx, string te, string ba)            
        {
                            XML.WriteStartElement("Elemento");
                                                        XML.WriteAttributeString("Tema"   , te);
                                                        XML.WriteAttributeString("Bairro" , ba);
                                                        XML.WriteAttributeString("Objeto" , ob);
                                                        XML.WriteAttributeString("Local"  , xy);
                                     if (ob == "Texto") XML.WriteAttributeString("Texto"  , tx);
                            XML.WriteEndElement();
        }
        public void  Rio_Fecha_Bairros ( List<XmlWriter> Lista_xml )                                                       
        { 
                            for (int t = 0; t < Lista_xml.Count; t++)
                                {
                                       XmlWriter XML = Lista_xml[t];
                                       XML.WriteEndElement(); //fecha bairros  
                                }
        }
        public void  Rio_Fecha_Cidade  ( List<XmlWriter> Lista_xml )                                                       
        { 
                            for (int t = 0; t < Lista_xml.Count; t++)
                                {
                                       XmlWriter XML  = Lista_xml[t];
                                       XML.WriteEndElement(); //fecha cidade 
                                       XML.Flush();
                                       XML.Close();
                                }
        }
        
        [CommandMethod("RIOKML")]
        public void  RIO_KML()                                                                                                                   
        {
                    XmlWriterSettings opc = new XmlWriterSettings();
                    opc.Indent            = true;
                    opc.CheckCharacters   = false;
                    List<XmlWriter> L_kml = new List<XmlWriter> { };
                    try
                    {
                          string       Pasta;
                          List<string> Temario;
                          List<string> Barrios;
                          Rio_Selecion_Tema(out Pasta, out Temario, out Barrios);
                          if (Temario.Count > 0)
                          {
                               for (int t = 0; t < Temario.Count; t++)
                               {
                                    string    Tema = Temario[t];
                                    string    ArqT = Pasta + "\\KML\\KML_Rio_" + Tema + ".KML";
                                    XmlWriter KML  = XmlWriter.Create ( ArqT , opc );
                                    Rio_KML_Cidade_Abrir ( KML , "Rio de Janeiro" , Tema );
                                    L_kml.Add( KML );                      
                               }
                               Rio_KML_Explorar_DWG( Pasta , Temario , L_kml , Barrios); 
                          }
                          Rio_KML_Cidade_Fecha( L_kml ); //Fecha a cidade para cada tema
                    }
                    catch   { }
                    finally { }
                    Mensaje( "Temas extraidos em arquivos KML.");
        }
        public void  Rio_KML_Cidade_Abrir  ( XmlWriter KML , string cidade , string te )                                                         
        {
                     // Valores de opacidade em Hex
                     //  0% = 00    70% = 46
                     // 20% = 14    80% = 50    
                     // 40% = 28    90% = 5a 
                     // 50% = 32   100% = FF            
                     string  co   =  Rio_Dados_Tematica(te)[2];
                     string  de   =  Rio_Dados_Tematica(te)[4];
                     string  Red  =  co.Substring(0, 2);
                     string  Gre  =  co.Substring(2, 2);
                     string  Blu  =  co.Substring(4, 2);
                     string  Alf  =  "9F";
                     string  pla  =  Alf + Blu + Gre + Red;
                     string  lin  =  "FFA9A9A9";
             
                     KML.WriteStartElement ( "kml" );
                         KML.WriteStartElement ( "Document" );
                             Rio_KML_Estilo_POL ( KML , "Poligo" + te , lin , "0.5" , pla );
                             Rio_KML_Estilo_PON ( KML , "Pontos" + te );
                             KML.WriteStartElement ( "Folder" );
                             KML.WriteStartElement ( "name"   );     KML.WriteString ( cidade + "_" + te );        KML.WriteEndElement();
                             KML.WriteStartElement ( "SimpleData" ); KML.WriteAttributeString( "Descrição" , de ); KML.WriteEndElement();
        }
        public void  Rio_KML_Processa_Bar  ( string pasta  , List<string> Temas, List<XmlWriter> LKML , string ArqDWG )                          
        {
                     string   A  = pasta +   "\\"  + ArqDWG;
                     string[] B  = Rio_Dados_Bairro( ArqDWG ).ToArray();
                     Database C  = new Database(false, true);

                     using (C)
                     {
                           C.ReadDwgFile( A , FileOpenMode.OpenForReadAndWriteNoShare, true, "");
                           Transaction tr = C.TransactionManager.StartOpenCloseTransaction();
                           using (tr)
                           {
                                 BlockTable       bt   = (BlockTable)tr.GetObject( C.BlockTableId, OpenMode.ForRead );
                                 BlockTableRecord btr  = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);
                                 foreach (ObjectId objId in btr)
                                 {
                                      Entity   en  =  (Entity)tr.GetObject(objId, OpenMode.ForRead);
                                      string   la  =  en.Layer.Substring (  0,  3  );
                                      string   to  =  Tipo_de_Objeto     (  en     );
                                      string   tx  =  Texto_do_Objeto    (  en     );
                                      string   xy  =  Extrair_Pontos_KML (  en, tr );
                                      string   id  =  en.Handle.Value.ToString();
                                      string   pi  =  ElevacaoPonto;
                                      string   ps  =  ElevacaoTotal;
                                      for (int t = 0; t < Temas.Count; t++)
                                      {
                                           XmlWriter KML  = LKML[t];                     // arquivo KML do tema
                                           string    tema = Temas[t];
                                           string    ex   = Rio_Dados_Tematica(tema)[3]; // valor de extrução do tema
                                           string    de   = Rio_Dados_Tematica(tema)[4]; // descrição do tema

                                           if (to != "Desconocido" && la.Contains(tema))
                                           {
                                                switch (to)
                                                {
                                                     case "Ponto":   Rio_KLM_Placemark_PNT( KML , tema , B , id , de , xy , ps      ); break;
                                                     case "Bloco":   Rio_KLM_Placemark_PNT( KML , tema , B , id , de , xy , ps      ); break;
                                                     case "Texto":   Rio_KLM_Placemark_PNT( KML , tema , B , id , de , xy , ps      ); break;
                                                     case "Linha":   Rio_KLM_Placemark_PNT( KML , tema , B , id , de , xy , ps      ); break;
                                                     case "Circulo": Rio_KLM_Placemark_PNT( KML , tema , B , id , de , xy , ps      ); break;
                                                     case "Poli_2D": Rio_KLM_Placemark_POL( KML , tema , B , id , de , xy , ps , ex ); break;
                                                     case "Poli_3D": Rio_KLM_Placemark_POL( KML , tema , B , id , de , xy , ps , ex ); break; 
                                                     default:                                                                          break;
                                                }
                                           } 
                                      } 
                                 }
                                 tr.Commit();
                           }
                           C.CloseInput( true );
                           C.Dispose();
                     } 
        } 
        public void  Rio_KML_Explorar_DWG  ( string pasta  , List<string> Temas, List<XmlWriter> LKML , List<string> Barrios)                    
        {
                            List<string> Arquivos = Rio_Montar_Barrios( Barrios );
                            try
                            {
                                        for (int a = 0; a < Arquivos.Count; a++)
                                        {
                                               string dwg = Arquivos[a];
                                               Rio_KML_Processa_Bar(pasta , Temas , LKML , dwg);
                                        }
                            }
                            catch   {  }
                            finally {  }

        }
        public void  Rio_KLM_Placemark_POL ( XmlWriter KML , string te , string[] B , string id , string de , string xy , string ps , string ex) 
        {
                                KML.WriteStartElement("Placemark"); 
                                    KML.WriteStartElement("name");        KML.WriteString( id            ); KML.WriteEndElement();
                                    KML.WriteStartElement("description"); KML.WriteString( te + " "  + de); KML.WriteEndElement();
                                    KML.WriteStartElement("styleUrl");    KML.WriteString( "#Poligo" + te); KML.WriteEndElement();
                                    KML.WriteStartElement("visibility");  KML.WriteString( "1"           ); KML.WriteEndElement();
                                    KML.WriteStartElement("ExtendedData");
                                         Rio_KML_Dado_Add ( KML , "Bairro_AP",     B[0]);
                                         Rio_KML_Dado_Add ( KML , "Bairro_RA",     B[1]);
                                         Rio_KML_Dado_Add ( KML , "Bairro_Nome",   B[3]);
                                         Rio_KML_Dado_Add ( KML , "Bairro_Zona",   B[4]);
                                         Rio_KML_Dado_Add ( KML , "Bairro_Codigo", B[5]);
                                         Rio_KML_Dado_Add ( KML , "Altura",          ps);
                                    KML.WriteEndElement ( );
                                    KML.WriteStartElement("Polygon");
                                       KML.WriteStartElement("extrude"   );     KML.WriteString( ex );        KML.WriteEndElement();
                                       KML.WriteStartElement("altitudeMode");   KML.WriteString("absolute");  KML.WriteEndElement();
                                       KML.WriteStartElement("outerBoundaryIs");
                                         KML.WriteStartElement("LinearRing");
                                           KML.WriteStartElement("coordinates"); KML.WriteString(xy); KML.WriteEndElement(); //Abre fecha Coordenadas
                                         KML.WriteEndElement();
                                       KML.WriteEndElement();
                                    KML.WriteEndElement();
                                KML.WriteEndElement();  //Fecha Placemark
        } 
        public void  Rio_KLM_Placemark_PNT ( XmlWriter KML , string te , string[] B , string id , string de , string xy , string ps)             
        {
                                    KML.WriteStartElement("Placemark"); 
                                        KML.WriteStartElement("name");              KML.WriteString(  id           ); KML.WriteEndElement();
                                        KML.WriteStartElement("description");       KML.WriteString(  te + " " + de); KML.WriteEndElement();
                                        KML.WriteStartElement("styleUrl");          KML.WriteString( "#Pontos" + te); KML.WriteEndElement();
                                        KML.WriteStartElement("ExtendedData");
                                            Rio_KML_Dado_Add ( KML , "Bairro_AP",     B[0]);
                                            Rio_KML_Dado_Add ( KML , "Bairro_RA",     B[1]);
                                            Rio_KML_Dado_Add ( KML , "Bairro_Nome",   B[3]);
                                            Rio_KML_Dado_Add ( KML , "Bairro_Zona",   B[4]);
                                            Rio_KML_Dado_Add ( KML , "Bairro_Codigo", B[5]);
                                            Rio_KML_Dado_Add ( KML , "Altura",          ps);
                                        KML.WriteEndElement();
                                        KML.WriteStartElement("Point");
                                          KML.WriteStartElement("coordinates");KML.WriteString (xy); KML.WriteEndElement(); //Abre fecha Coordenadas
                                        KML.WriteEndElement(); //Fecha Point
                                KML.WriteEndElement();        //Fecha Placemark

        }
        public void  Rio_KML_Dado_Add      ( XmlWriter KML , string da , string val)                                                             
        {
                        KML.WriteStartElement("Data");
                            KML.WriteAttributeString("name", da);
                                KML.WriteStartElement("displayName"); KML.WriteString( da  ); KML.WriteEndElement();
                                KML.WriteStartElement("value"      ); KML.WriteString( val ); KML.WriteEndElement();
                        KML.WriteEndElement();
        }
        public void  Rio_KML_Estilo_POL    ( XmlWriter KML , string id , string lin , string esp , string pla )                                  
        {  
                               KML.WriteStartElement("Style"); KML.WriteAttributeString( "id" , id);
                                        KML.WriteStartElement("LineStyle");
                                             KML.WriteStartElement("width"); KML.WriteString( esp ); KML.WriteEndElement();
                                             KML.WriteStartElement("color"); KML.WriteString( lin ); KML.WriteEndElement();
                                        KML.WriteEndElement();

                                        KML.WriteStartElement("PolyStyle");
                                             KML.WriteStartElement("color"); KML.WriteString( pla ); KML.WriteEndElement();
                                        KML.WriteEndElement();
                               KML.WriteEndElement();
        }
        public void  Rio_KML_Estilo_PON    ( XmlWriter KML , string id                                        )                                  
        {  
                                   KML.WriteStartElement("Style");  KML.WriteAttributeString( "id" , id );
                                        KML.WriteStartElement("IconStyle");
                                            KML.WriteStartElement("Icon");
                                                KML.WriteStartElement("href");
                                                    KML.WriteString("Ponto.png");
                                                KML.WriteEndElement();
                                            KML.WriteEndElement();
                                        KML.WriteEndElement();
                                   KML.WriteEndElement();
        }
        public void  Rio_KML_Cidade_Fecha  ( List<XmlWriter> Lxml )                                                                              
        {
                            for (int t = 0; t < Lxml.Count; t++)
                            {
                                 XmlWriter XML = Lxml[t];
                                 XML.WriteEndElement(); //fecha Folder 
                                 XML.WriteEndElement(); //fecha Document 
                                 XML.WriteEndElement(); //fecha kml 
                                 XML.Flush();
                                 XML.Close();
                            }
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
