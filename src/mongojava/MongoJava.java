package mongojava;
import com.mongodb.*;
import com.mongodb.client.MongoCollection;
import com.mongodb.client.MongoDatabase;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.bson.Document;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class MongoJava {

    static MongoDatabase db;
    static DBCollection coleccion;
    
    public static void agregarDatos() {
        
        /*
         * Se crea una coleccion con el nombre desado, en el caso de que esta no
         * exista se crea, si no se trabaja en ella.
         */
        MongoCollection<Document> collection = db.getCollection("EstoEsUnaPrueba"); 
        
        try {
            String excelPath = "archivofinal.xlsx";
            FileInputStream inputStream = new FileInputStream(new File(excelPath));
            XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
            Sheet firstSheet = workbook.getSheetAt(0);
            Iterator<Row> iterator = firstSheet.iterator();
            
            while (iterator.hasNext()){
                Row nextRow = iterator.next();
                /*
                *Se salta a la segunda columna porque la primera tiene los siguientes atributos
                */
                Iterator<Cell> cellIterator = nextRow.cellIterator();
                while(cellIterator.hasNext()){
                    Document document = new Document("tittle", "IPGH");
                    for(int i = 1; i < 13; i++){
                        Cell cell = cellIterator.next();
                        switch(i){
                            case 1:
                                document.append("NombrePeriodico", cell.toString());
                                break;
                            case 2:
                                document.append("Número", cell.toString());
                                break;
                            case 3:
                                document.append("Fecha", cell.toString());
                                break;
                            case 4:
                                document.append("Evento", cell.toString());
                                break;
                            case 5:
                                document.append("Ubicacion", cell.toString());
                                break;
                            case 6:
                                document.append("Latitud", cell.toString());
                                break;
                            case 7:
                                document.append("Longitud", cell.toString());
                                break;
                            case 8:
                                document.append("VelocidadViento", cell.toString());
                                break;
                            case 9:
                                document.append("Creciente", cell.toString());
                                break;
                            case 10:
                                document.append("Temperatura", cell.toString());
                                break;
                            case 11:
                                document.append("EfectoDetectado", cell.toString());
                                break;
                            case 12:
                                document.append("URL", cell.toString());
                                break;
                        }
                    }
                    collection.insertOne(document);
                }
            }
            System.out.println("Documento insertado con exito :D");
            workbook.close();
            inputStream.close();
        } catch (IOException ex) {
            Logger.getLogger(MongoJava.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
    
    public static void evitarLogs(){
        /*
         * Este metodo previene que salgan logs de creaciones, los cuales solo
         * distraen en la consola.
         */
        Logger mongoLogger = Logger.getLogger( "org.mongodb.driver" );
        mongoLogger.setLevel(Level.SEVERE);
    }
    
    public static void crearConexion() {
        evitarLogs();
        /*
         *Aqui se hace un cliente para acceder y se crea una bd si no existe
         */
        MongoClient mongo = new MongoClient("localhost", 27017);
        db = mongo.getDatabase("prueba1");
        System.out.println("Conexion realizada con exito");
    }
    
    public static void main(String[] args) {
        crearConexion();
        agregarDatos();
    }
}