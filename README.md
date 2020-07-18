# Rio
O projeto Rio trata de um sistema programado para extrair e re-estruturar dados em formato XML dos arquivos DWG da cidade de Rio de Janeiro. 
O projeto objetiva ter um repositório de dados que permita integrar bases de informação existentes em DWG com aplicações BIM direcionadas para o projeto urbano. 

Os arquivos da série Rio_KML são estruturados para que possam ser abertos diretamente em ambientes GIS como o QGIS ou o GoogleEarth. 
As coordenadas estão em formato Latitude e Longitude decimal.
Exemplo de um elemento:
<Placemark>
      <name>5600</name>
      <description>Prédios</description>
      <visibility>1</visibility>
      <Bairro_AP>2.0</Bairro_AP>
      <Bairro_RA>V</Bairro_RA>
      <Bairro_Nome>Copacabana</Bairro_Nome>
      <Bairro_Zona>Copacabana</Bairro_Zona>
      <Bairro_Codigo>024</Bairro_Codigo>
      <Polygon>
        <tessellate>1</tessellate>
        <outerBoundaryIs>
          <LinearRing>
            <coordinates> -43.1769138575802,-22.9604190879071,0 -43.1770549735727,-22.9605928277586,0 -43.1771148217001,-22.9605511637856,0 -43.1769737056663,-22.9603774239845,0</coordinates>
          </LinearRing>
        </outerBoundaryIs>
      </Polygon>
    </Placemark>
    
