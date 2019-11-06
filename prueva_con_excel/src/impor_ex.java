
import com.itextpdf.layout.property.*;
//import org.apache.poi.ss.usermodel.HorizontalAlignment;
/*import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;*/
import com.itextpdf.io.font.FontConstants;
import com.itextpdf.io.image.ImageDataFactory;
import com.itextpdf.kernel.font.PdfFont;
import com.itextpdf.kernel.font.PdfFontFactory;
import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.Border;
import com.itextpdf.layout.element.Image;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
//import com.itextpdf.layout.property.HorizontalAlignment;
import com.itextpdf.layout.property.TextAlignment;
import com.itextpdf.layout.property.VerticalAlignment;
import java.io.File;
import java.io.FileInputStream;
//import java.io.IOException;
//import java.io.FileOutputStream;
import java.io.IOException;
//import java.io.OutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.DecimalFormat;
//import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.GregorianCalendar;


import javax.swing.JFileChooser;
import javax.swing.JOptionPane;
//import javax.swing.UIManager;
//import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;


import java.util.Iterator;
import java.util.Scanner;
import java.util.StringTokenizer;
import java.util.logging.Level;
import java.util.logging.Logger;
//import javax.swing.ImageIcon;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
/*import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumnModel;*/
import org.apache.poi.ss.usermodel.Cell;
//import org.apache.poi.ss.usermodel.CellType;
//import static org.apache.poi.ss.usermodel.CellType.BLANK;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;
//import static org.apache.poi.ss.usermodel.CellType.STRING;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class impor_ex extends javax.swing.JFrame {
 
    DefaultTableModel modelo; //declara el modeo de la tabla
    //java.util.Date fecha = new Date();
    boolean bandera= false; // bandera para la verifacion de carga de archivos
    int comprobar= 0;// bandera auxiliar para la verificacion de carga de archivos
    double suma_Global[]=new double[80];//suma total de las califcaciones por alumno
    double promedio_Global[]=new double[80];//promedio por cada uno de los alumnos
    String arrayA[]= new String[25]; //obtendra el nombre las columnas
    String matricula[]= new String[80];//arreglo que tiene todas las matriculas
    
            
    public impor_ex() {     
        initComponents();
        this.setSize(905, 700);
        // this.setSize(Toolkit.getDefaultToolkit().getScreen Size());
        this.setLocationRelativeTo(null);//coloca el Frame a la mitad de la pantalla
        //Las siguientes 2 lineas colocan en automatico la fecha del dia en el JDateChooser
        Calendar c2 = new GregorianCalendar();
        calendario.setCalendar(c2);  
        //seleccion.setEnabled(false);//esta funcion activa o desactiva el componente
        seleccion.setSelected(true);
    }
    
    
    public String[] fecha(){
        
        Date fechaCal = calendario.getDate();
        String f_form[]= new String[3];       
        
        SimpleDateFormat Formato = new SimpleDateFormat("d");
        SimpleDateFormat FormatoMes = new SimpleDateFormat("MM");
        SimpleDateFormat FormatoAnnio = new SimpleDateFormat("yyyy");
        
        f_form[0]=Formato.format(fechaCal.getTime());
        f_form[1]=FormatoMes.format(fechaCal.getTime());
        f_form[2]=FormatoAnnio.format(fechaCal.getTime());
        
        long seleccionado_mes = Integer.parseInt(f_form[1]);
        
        if(seleccionado_mes==1){
            f_form[1]="ENERO";
        }if(seleccionado_mes==2){
            f_form[1]="FEBRERO";
        }if(seleccionado_mes==3){
            f_form[1]="MARZO";
        }if(seleccionado_mes==4){
            f_form[1]="ABRIL";
        }if(seleccionado_mes==5){
            f_form[1]="MAYO";
        }if(seleccionado_mes==6){
            f_form[1]="JUNIO";
        }if(seleccionado_mes==7){
            f_form[1]="JULIO";
        }if(seleccionado_mes==8){
            f_form[1]="AGOSTO";
        }if(seleccionado_mes==9){
            f_form[1]="SEPTIEMBRE";    
        }if(seleccionado_mes==10){
            f_form[1]="OCTUBRE";
        }if(seleccionado_mes==11){
            f_form[1]="NOVIEMBRE";
        }if(seleccionado_mes==12){
            f_form[1]="DICIEMBRE";
        }     
        return f_form;
    }
    
    
    
    public void crear_tabla(String cadena_path, boolean chec){
       
       
              
	String hoja = "Hoja1";

        int f=0;//contador de filas
        int c=0;//contador de celdas
        int cal=0;//contador de las calificaciones
    
	//try (FileInputStream file = new FileInputStream(new File(rutaArchivo))) {
        try (FileInputStream file = new FileInputStream(new File(cadena_path))) {
	    // leer archivo excel
            
            XSSFWorkbook worbook = new XSSFWorkbook(file);
            //obtener la hoja que se va leer
            XSSFSheet sheet = worbook.getSheetAt(0);
            //obtener todas las filas de la hoja excel
            Iterator<Row> rowIterator = sheet.iterator();
   
            Row row, row2;
            
            //While's que recorren todas las filas y todas las celdad del Excel
            //para tenet un contador 'f= num_filas' y 'c=num_celdas'
            while ( rowIterator.hasNext()) { 
                row2 = rowIterator.next();
                Iterator<Cell> cellIterator = row2.cellIterator();
                Cell cellCon;
                c=0;
                while ( cellIterator.hasNext()) { 
                   cellCon = cellIterator.next();
                   c++;//aumenta el contador de celdas el excel
                } 
                f++;//aumenta el contador de filas el excel
            }

            
            //*****************comprobar que en los numeros no haya otros caracteres 
            //                 ni numeros negativos
            int bp=0;//bandera de comprobacion
            for(int fv=1; fv<f; fv++){
                for(int cv=3; cv<c; cv++){
                    row2=sheet.getRow(fv);//recorre las filas en el indice 'fc'
                    Cell cell = row2.getCell(cv);//recorre las columnas en el indice 'fc'
                    if( cell.getCellType() != NUMERIC || cell.getNumericCellValue()< 0 ||  cell.getNumericCellValue()> 10){
                        bp++;
                    }   
                }
            }
            
            if(bp>0){
                JOptionPane.showMessageDialog(null,"EN LA PARTE DE DE NUMEROS PUEDE HABER:\n "
                                              + "--> UN CARACTER DISTINTO\n --> NEGATIVO\n -->MAYOR A 10");                
            }else{
            //*****************comprobar que en los numeros no haya otros caracteres

                //genera un modelo de tabla 
                modelo =(DefaultTableModel) tabla.getModel();
            
                //objeto que contendra el nombre de las materias
                Object[] nombres_columnas= new Object[c+1];
                nombres_columnas[0]="Acepta";
                        
           
            
                // Esta validacion funciona para saber si es la primera vez que se carga
                // un archivo a la tabla si es la segunda vez o mas se elimina el
                // contenido de la tabla  
                if(chec==true){     
                    int a =modelo.getRowCount()-1;//se cuenta el numero de filas
                    //for que elimina las filas de la tabla
                    for(int i=a; i>=0; i--){
                        modelo.removeRow(i);//elimina las filas
                    }
                
                    int filas_t = modelo.getRowCount();//cuenta las filas de la tabla
                
                    //este for asigna 0 al arreglo que tiene las sumas y promedios 
                    //de calificaiones
                    for(int j=0; j<filas_t; j++){
                        suma_Global[j]=0;
                        promedio_Global[j]=0;
                        arrayA[j]="";
                        matricula[j]="";
                    }
                }                    
            
                //for que obtiene el nombre de la primera fila del Excel, 
                // para asignarlo como nombre de columnas en la tabla
                for(int rc=0; rc<c; rc++){ 
                    row=sheet.getRow(0);//solo maneja la primera fila
                    Cell cell = row.getCell(rc);//obtine el valor de cada celda de la fila
                    arrayA[rc]=cell.toString(); //se asigna y onvierte a cadena en arrayA[]  
                    nombres_columnas[rc+1]=arrayA[rc];//se llena el arreglo con los nombres de materias
                } 
            
                //le asgina el nombre "NUM" a la segunda columna
                nombres_columnas[1]="NUM";
             
                //objeto que contendra todos los datos 
                Object[] data= new Object[80]; 
            
                //le coloca el nombre columnas en la tabla
                modelo.setColumnIdentifiers(nombres_columnas);
                      
                System.out.println("\n\nNumero de Filas: "+f+ "\nNumero de Columnas: "+c+"\n");
            
            
                //arreglo que convirte en cadena los datos de 'data[]'      
                String arr[]= new String[c];   
            
                //arreglo que contiene las sumas de todas las calificaciones por alumno
                double suma[]=new double[f-1];
                double promedio[]=new double[f-1];
            
                //for´s que recorren todas las filas y todos las celdas
                for(int rf=1; rf<f; rf++){

                    data[0]=true;
                    for(int rc=0; rc<c; rc++){ 
                    
                        row=sheet.getRow(rf);//recorre las filas en el indice 'rf'
                        Cell cell = row.getCell(rc);//recorre las celdas en el indice 'rc'
                    
                        if(cell.getCellType()==NUMERIC){                      
                            //Redondea las calificaciones del excel a el entero mas proximo
                            arr[rc]=String.valueOf(Math.round(cell.getNumericCellValue()));
                        
                        }else{
                            arr[rc]=cell.toString(); //convierte la celda en cadena  
                        }
                    
                    
                        data[rc+1]=arr[rc];  //asigna los valores de arr[] a data[]
                    
                        //se empieza a partir de la celda 2 porque las calificaciones empiezan en esa celda
                        if(rc>=3){ 
                            suma[rf-1] = suma[rf-1] + Math.round(cell.getNumericCellValue()); //hace la suma de las calficaciones
                            promedio[rf-1] = suma[rf-1]/(c-3);  //calcula el promedio     
                        }
                    }
                    modelo.addRow(data);//se agregan los datos al modelo    
                }
            
                // Este for guarda la suma de las calificaciones y el promedio 
                // en el arreglo "_Goblal" con el contador 'cal' que tiene como
                //proposito asignar el numero de elementos totales asignados antes
                for(int t=0; t<(f-1); t++){ 
                    suma_Global[t]=suma[t];
                    promedio_Global[t]=promedio[t];    
                }
            }
           
	} catch (Exception e) {
            e.getMessage();
            JOptionPane.showMessageDialog(null, "Hay ERROR en las celdas");        
	}       
    }

    /**
     * This method is called from within the constructor to initialize the form.
     * WARNING: Do NOT modify this code. The content of this method is always
     * regenerated by the Form Editor.
     */
    @SuppressWarnings("unchecked")
    // <editor-fold defaultstate="collapsed" desc="Generated Code">//GEN-BEGIN:initComponents
    private void initComponents() {

        recuadros_Seleccion = new javax.swing.ButtonGroup();
        jLabel1 = new javax.swing.JLabel();
        btn_Importar = new javax.swing.JButton();
        jScrollPane1 = new javax.swing.JScrollPane();
        tabla = new javax.swing.JTable();
        imprimir = new javax.swing.JButton();
        semestre = new javax.swing.JComboBox<>();
        jLabel2 = new javax.swing.JLabel();
        jLabel3 = new javax.swing.JLabel();
        grupo = new javax.swing.JComboBox<>();
        salir = new javax.swing.JButton();
        jLabel4 = new javax.swing.JLabel();
        calendario = new com.toedter.calendar.JDateChooser();
        seleccion = new javax.swing.JRadioButton();
        deseleccion = new javax.swing.JRadioButton();

        setDefaultCloseOperation(javax.swing.WindowConstants.EXIT_ON_CLOSE);
        getContentPane().setLayout(null);

        jLabel1.setFont(new java.awt.Font("Yu Gothic", 0, 24)); // NOI18N
        jLabel1.setHorizontalAlignment(javax.swing.SwingConstants.CENTER);
        jLabel1.setText("Escoge tu Documento");
        getContentPane().add(jLabel1);
        jLabel1.setBounds(270, 10, 350, 50);

        btn_Importar.setFont(new java.awt.Font("Yu Gothic", 0, 14)); // NOI18N
        btn_Importar.setText("Subir Archivo");
        btn_Importar.setBorder(new javax.swing.border.SoftBevelBorder(javax.swing.border.BevelBorder.RAISED, new java.awt.Color(153, 204, 255), new java.awt.Color(153, 204, 255), new java.awt.Color(153, 204, 255), new java.awt.Color(153, 204, 255)));
        btn_Importar.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                btn_ImportarActionPerformed(evt);
            }
        });
        getContentPane().add(btn_Importar);
        btn_Importar.setBounds(40, 90, 140, 40);

        tabla.setModel(new javax.swing.table.DefaultTableModel(
            new Object [][] {

            },
            new String [] {
                "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""
            }
        ) {
            Class[] types = new Class [] {
                java.lang.Boolean.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class, java.lang.Object.class
            };

            public Class getColumnClass(int columnIndex) {
                return types [columnIndex];
            }
        });
        jScrollPane1.setViewportView(tabla);

        getContentPane().add(jScrollPane1);
        jScrollPane1.setBounds(30, 250, 820, 250);

        imprimir.setFont(new java.awt.Font("Yu Gothic", 0, 14)); // NOI18N
        imprimir.setText("Imprimir");
        imprimir.setBorder(new javax.swing.border.SoftBevelBorder(javax.swing.border.BevelBorder.RAISED, new java.awt.Color(153, 255, 153), new java.awt.Color(153, 255, 153), new java.awt.Color(153, 255, 153), new java.awt.Color(153, 255, 153)));
        imprimir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                imprimirActionPerformed(evt);
            }
        });
        getContentPane().add(imprimir);
        imprimir.setBounds(580, 580, 150, 40);

        semestre.setFont(new java.awt.Font("Yu Gothic", 0, 12)); // NOI18N
        semestre.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Seleccionar", "PRIMER SEMESTRE", "SEGUNDO SEMESTRE", "TERCER SEMESTRE", "CUARTO SEMESTRE", "QUINTO SEMESTRE", "SEXTO SEMESTRE" }));
        getContentPane().add(semestre);
        semestre.setBounds(260, 180, 150, 30);

        jLabel2.setFont(new java.awt.Font("Yu Gothic", 0, 12)); // NOI18N
        jLabel2.setText("Semestre");
        getContentPane().add(jLabel2);
        jLabel2.setBounds(200, 180, 60, 30);

        jLabel3.setText("Grupo");
        getContentPane().add(jLabel3);
        jLabel3.setBounds(460, 180, 50, 30);

        grupo.setFont(new java.awt.Font("Yu Gothic", 0, 12)); // NOI18N
        grupo.setModel(new javax.swing.DefaultComboBoxModel<>(new String[] { "Seleccionar", "1", "2" }));
        getContentPane().add(grupo);
        grupo.setBounds(510, 180, 110, 30);

        salir.setFont(new java.awt.Font("Yu Gothic", 0, 14)); // NOI18N
        salir.setText("Salir");
        salir.setBorder(new javax.swing.border.SoftBevelBorder(javax.swing.border.BevelBorder.RAISED, new java.awt.Color(255, 0, 0), new java.awt.Color(255, 0, 0), new java.awt.Color(255, 0, 0), new java.awt.Color(255, 0, 0)));
        salir.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                salirActionPerformed(evt);
            }
        });
        getContentPane().add(salir);
        salir.setBounds(130, 580, 120, 40);

        jLabel4.setText("FECHA");
        getContentPane().add(jLabel4);
        jLabel4.setBounds(620, 90, 60, 30);
        getContentPane().add(calendario);
        calendario.setBounds(670, 90, 170, 30);

        recuadros_Seleccion.add(seleccion);
        seleccion.setText("Seleccionar Todas");
        seleccion.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                seleccionActionPerformed(evt);
            }
        });
        getContentPane().add(seleccion);
        seleccion.setBounds(50, 510, 120, 30);

        recuadros_Seleccion.add(deseleccion);
        deseleccion.setText("Deseleccionar Todo");
        deseleccion.addActionListener(new java.awt.event.ActionListener() {
            public void actionPerformed(java.awt.event.ActionEvent evt) {
                deseleccionActionPerformed(evt);
            }
        });
        getContentPane().add(deseleccion);
        deseleccion.setBounds(170, 510, 140, 30);

        pack();
    }// </editor-fold>//GEN-END:initComponents

    private void btn_ImportarActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_btn_ImportarActionPerformed
        JFileChooser fileChooser = new JFileChooser();
        FileNameExtensionFilter filter = new FileNameExtensionFilter("Archivos Excel (*.xls)", "xlsx");//filtro de busqueda
        fileChooser.setAcceptAllFileFilterUsed(false);
        fileChooser.setFileFilter(filter);  //se carga el filtro 
        fileChooser.setDialogTitle("BUSCAR ARCHIVO"); //nombre de ventana 
        
        int seleccion=fileChooser.showOpenDialog(this);// 'seleccion' es el numero de opcion 
                                                       // que se escoge el usuario
        //Si el usuario, pincha en aceptar
	if(seleccion==JFileChooser.APPROVE_OPTION){
            //
            if(comprobar==0){
                bandera=false;
                comprobar=1;
            }else{
                bandera=true;
            }
   
            File archivo = fileChooser.getSelectedFile();//obtiene la direccion path del archivo
            String path_archivo=archivo.getAbsolutePath();//se convierte la direc. path en String
  
            //System.out.println("archivo seleccionado: "+ path_archivo);
            crear_tabla(path_archivo, bandera);//se llama a la funcion de crear tabla
   
	}	
        
    }//GEN-LAST:event_btn_ImportarActionPerformed

    private void imprimirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_imprimirActionPerformed

        
        System.out.println("\n"); 
        String arrfecha[]=fecha();
        for(int y=0; y<3; y++){
            System.out.println(arrfecha[y]);
        }
        String path = null;
        JFileChooser file = new JFileChooser();
        file.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        file.setMultiSelectionEnabled(false);
        file.setDialogTitle("GUARDAR ARCHIVOS"); //nombre de ventana 
        int seleccion=file.showOpenDialog(this);// 'seleccion' es el numero de opcion 
                                                       // que se escoge el usuario
        //Si el usuario, pincha en aceptar
        if(seleccion==JFileChooser.APPROVE_OPTION){
            //
            /*if(comprobar==0){
                bandera=false;
                comprobar=1;
            }else{
                bandera=true;
            }*/
            path = file.getSelectedFile().getAbsolutePath();
             //opcion del 1 al 2 extraen la opcion dentro del jcombobox
        String opcion1 = semestre.getSelectedItem().toString();
        String opcion2 = grupo.getSelectedItem().toString();
               
        // Si la opcion NO es seleccionar se arroja un mensaje que advierte que 
        // opcion esta sin seleccionar
        if(opcion1.equals("Seleccionar")){
            JOptionPane.showMessageDialog(null, "No selecciono: SEMESTRE");
        }
        
        if(opcion2.equals("Seleccionar")){
            JOptionPane.showMessageDialog(null, "No selecciono: GRUPO");
        }              
        
        //en caso de que ninguna opcion quede sin seleccionar se prodece a mandar/imprimir los datos 
        if(opcion1.equals("Seleccionar") || opcion2.equals("Seleccionar")  ){
        }else{
  
            System.out.println("\n\nEl Semestre es: "+opcion1);
            System.out.println("El Grupo es: "+opcion2 + "\n");
            
            
            int fil = modelo.getRowCount();//cuenta el numero de filas
            DecimalFormat df = new DecimalFormat("#.0");//da el formato de impresion a 1 decima 
                for(int e=0; e<fil; e++){
                    //Imprime todas las calificaciones de sumas y promedio 
                    System.out.println("la SUMA FINAL: "+df.format(suma_Global[e])+ "   El PROMEDIO FINAL: "+df.format(promedio_Global[e]));
                }
            
            System.out.print("\n");
            int cols = modelo.getColumnCount();//se cuenta el numero de columnas
            int fils = modelo.getRowCount();//cuenta el numero de filas

            
       // SEGCION FINAL PARA IMPRIMIR LOS DATOS                                                                               
            
                                                                                        
            
            //Este for sirve para obtener el valor en String de la tabla 
            // y guarda el dato en el arreglo arrFinal[]
            for(int i=0; i<fils; i++) {
                
                //Este arreglo contiene los datos de la fila ya para imprimir
                String arrFinal[]=new  String[25];
                
                //verifica si esta seleccionado oh no el check box
                String compara = modelo.getValueAt(i, 0).toString();
                
                if(compara.equals("true")){
                    for(int j=3; j<cols; j++){
                        String cont_tabla=modelo.getValueAt(i,j).toString();
                        
                        //valor_matricula contiene el valor de la matricula manda este valor al pdf
      /*WWWWWWWWWWWWWW*/String valor_matricula=modelo.getValueAt(i,2).toString();/*WWWWWWWWWWWWWW*/
                        
                        
                        
                        arrFinal[j-3]=cont_tabla;  
                        //esta funcion redondea el promedio a solo una decima y  guarda en redondeado
                        double redondeado = new BigDecimal(promedio_Global[i]).setScale(1, RoundingMode.HALF_EVEN).doubleValue();
                        
                       
                        System.out.print(arrFinal[j-3]+" "  );      
                        
                        try {
                            //****************************************************************************
                            //*             De   Aqui  mandar los datos al codigo de PDF                 *
                            //****************************************************************************

                            crearPdf(arrFinal,opcion1,opcion2,redondeado,cols-1,path,valor_matricula,arrfecha);
                        } catch (IOException ex) {
                            Logger.getLogger(impor_ex.class.getName()).log(Level.SEVERE, null, ex);
                        }
                    }

                    System.out.print("\n");  
                }//if
            }//for de las filas
            
            //con este for se Extrae la fecha
            
        }     
            
        }else{
            JOptionPane.showMessageDialog(null, "guardado cancelado");
        }
          
    }//GEN-LAST:event_imprimirActionPerformed

    public void crearPdf(String[] nombre, String grado, String grupo, double redondeo, int col, String ruta,String matri,String[] fech) throws IOException {
        Scanner lee = new Scanner(System.in);

        	
        
        String dest = ruta+"\\boletas_"+nombre[0]+"_"+grado+""+grupo+".pdf";
        PdfWriter writer = new PdfWriter(dest);//escribe el pdf en la ruta que se le asigna 
        PdfDocument pdf = new PdfDocument(writer);
        try (Document documento = new Document(pdf, PageSize.Default)) {
           
            Image escu = new Image(ImageDataFactory.create("src\\imagenes\\escudo edo.png"));
            escu.scaleAbsolute(130,60);
            //escu.setTextAlignment(TextAlignment.LEFT);
            Image logo = new Image(ImageDataFactory.create("src\\imagenes\\logo final CBT.jpg"));          //pasamos la ruta a imageDataFactoryy devuelve un objeto y manar informacion de la imagen que itext puede leer
            logo.scaleAbsolute(60, 60);
            float suma = 0;
            int con = 0;
            //logo.setTextAlignment(TextAlignment.RIGHT);

            //escu.setAlignment();
            PdfFont font = PdfFontFactory.createFont(FontConstants.HELVETICA_BOLDOBLIQUE);

            PdfFont font1 = PdfFontFactory.createFont(FontConstants.COURIER);

            Table tabla1 = new Table(new float[]{7,7,7,7});
            tabla1.setWordSpacing(0);
            int cant;
            
            tabla1.addCell("MATERIA").addCell("CICLO").addCell("CALIFICACION").addCell("OBSERVACIONES");
            for(int i = 3 ;i < col;i++){
                
                    if(nombre[i-2]!=null){
                        if(Integer.parseInt(nombre[i-2]) >5){
                            tabla1.addCell(arrayA[i]).addCell("2019-2020").addCell(nombre[i-2]).addCell("APROVADA");
                           // System.out.println("i= "+i+" "+nombre[i]);
                        }else{
                            tabla1.addCell(arrayA[i]).addCell("2019-2020").addCell(nombre[i-2]).addCell("REPROVADA");
                           // System.out.println("i= "+i+" "+nombre[i]);
                        }
                    }
                    
                System.out.println("tamaño de arryA = "+arrayA[i]);
            }
          
            
            Paragraph p1 = new Paragraph()
                    .add(escu).setHorizontalAlignment(HorizontalAlignment.LEFT)
                    //setHorizontalAlignment(HorizontalAlignment.LEFT)
                    .add("                                                                               ")
                    .add(logo).setHorizontalAlignment(HorizontalAlignment.RIGHT);
            Paragraph p2 = new Paragraph()//.setFont(font)
                    .add("__________________________________________________________________").setFont(font).setTextAlignment(TextAlignment.CENTER)
                    .add("\n                    BOLETA DE CALIFICACIONES                    ").setFont(font).setTextAlignment(TextAlignment.CENTER);
            Paragraph p3 = new Paragraph()
                    .add("\nLA DIRECCION DE LA ESCUELA                                                        C.C.T 15ECT0166Q").setFont(font1).setFontSize(8f).setTextAlignment(TextAlignment.LEFT);

            Paragraph mat = new Paragraph()
                    .add(matri).setFont(font).setFontSize(10f).setTextAlignment(TextAlignment.CENTER);
            //documento.add(logo.setHorizontalAlignment(HorizontalAlignment.RIGHT));
            float promedio = suma / con;
            documento.add(p1.setTextAlignment(TextAlignment.CENTER));
            documento.add(p2);
            documento.add(p3);
            documento.add(new Paragraph("\nCBT No. 3, Zumpango").setFont(font).setFontSize(8f).setTextAlignment(TextAlignment.CENTER));
            documento.add(new Paragraph("\nESTABLECIDO EN ADOLFO LOPEZ MATEOS, ZUMPANGO").setFont(font1).setFontSize(8f).setTextAlignment(TextAlignment.LEFT));
            documento.add(new Paragraph("\nHACE CONSTAR QUE SEGUN REGISTROS QUE OBRAN EN EL ARCHIVO DE ESTE PLANTEL:").setFont(font1).setFontSize(8f).setTextAlignment(TextAlignment.LEFT));
            documento.add(new Paragraph(nombre[0]).setFont(font).setFontSize(10f).setTextAlignment(TextAlignment.CENTER));
            documento.add(new Paragraph("ES ALUMNO(A) DEL " + grado + " CON MATRICULA: ").setFont(font1).setFontSize(8f).setTextAlignment(TextAlignment.LEFT)).add(mat);
            documento.add(new Paragraph(" BACHILLERATO TECNOLOGICO CON LA CARRERA DE TECNICO EN INFORMATICA").setFont(font).setFontSize(10f).setTextAlignment(TextAlignment.CENTER));
            documento.add(new Paragraph("\n EN EL GRUPO         " + grupo + "        SUSTENTO LOS EXAMENES FINALES DE LAS MATERIAS QUE ACONTINUACION SE ANOTAN").setFont(font1).setFontSize(8f).setTextAlignment(TextAlignment.LEFT));
            documento.add(tabla1.setBorder(Border.NO_BORDER).setHorizontalAlignment(com.itextpdf.layout.property.HorizontalAlignment.CENTER).setFont(font).setFontSize(8f).setTextAlignment(TextAlignment.CENTER));
            documento.add(new Paragraph("PROMEDIO: " + redondeo).setFont(font).setFontSize(8f).setTextAlignment(TextAlignment.CENTER));
            documento.add(new Paragraph("LA CALIFICACION MINIMA APROBATORIA ES DE 6 (SEIS) PUNTOS").setFont(font).setFontSize(8f).setTextAlignment(TextAlignment.CENTER));
            documento.add(new Paragraph("ESTA BOLETA NO ES VALIDA SI PRESENTA BORRADURAS O ALTERACIONES").setFont(font1).setFontSize(8f).setTextAlignment(TextAlignment.CENTER).setVerticalAlignment(VerticalAlignment.TOP));
            documento.add(new Paragraph("\nZUMPANGO MEX., A LOS " + fech[0]+  " DIAS DEL MES DE " +fech[1]+ " DE "+fech[2] ).setFont(font).setFontSize(8f).setTextAlignment(TextAlignment.CENTER).setVerticalAlignment(VerticalAlignment.TOP));
            documento.add(new Paragraph("\nDIRECTO(A) ESCOLAR").setFont(font).setFontSize(8f).setTextAlignment(TextAlignment.CENTER).setVerticalAlignment(VerticalAlignment.BOTTOM));
            documento.add(new Paragraph("\n \n   "));
            documento.add(new Paragraph("__________________________________").setTextAlignment(TextAlignment.CENTER));
            documento.add(new Paragraph("MTRO. JUAN MANUEL LONGINOS CALLEJA").setFont(font).setFontSize(8f).setTextAlignment(TextAlignment.CENTER));
        } //pasamos la ruta a imageDataFactoryy devuelve un objeto y manar informacion de la imagen que itext puede leer //pasamos la ruta a imageDataFactoryy devuelve un objeto y manar informacion de la imagen que itext puede leer
    }
    
    public void process(Table tabla, String line, PdfFont font, boolean isHeader) {
        StringTokenizer token = new StringTokenizer(line, ",");
        while (token.hasMoreTokens()) {
            if (isHeader) {
                tabla.addHeaderCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(token.nextToken()).setFont(font)));
            } else {
                tabla.addCell(new com.itextpdf.layout.element.Cell().add(new Paragraph(token.nextToken()).setFont(font)));
            }
        }
    }
    
    private void salirActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_salirActionPerformed
        System.exit(0);
    }//GEN-LAST:event_salirActionPerformed

    private void seleccionActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_seleccionActionPerformed
        int fils = modelo.getRowCount();//cuenta el numero de filas
        
        //este for da el estado de verdadero a los check box
        for(int fg=0; fg<fils; fg++)         
                modelo.setValueAt(true, fg,0);         
    }//GEN-LAST:event_seleccionActionPerformed

    private void deseleccionActionPerformed(java.awt.event.ActionEvent evt) {//GEN-FIRST:event_deseleccionActionPerformed
        int fils = modelo.getRowCount();//cuenta el numero de filas
        
        //este for da el estado de falso a los check box
        for(int fg=0; fg<fils; fg++)
            modelo.setValueAt(false, fg,0);        
    }//GEN-LAST:event_deseleccionActionPerformed

    /**
     * @param args the command line arguments
     */
    public static void main(String args[]) {
        /* Set the Nimbus look and feel */
        //<editor-fold defaultstate="collapsed" desc=" Look and feel setting code (optional) ">
        /* If Nimbus (introduced in Java SE 6) is not available, stay with the default look and feel.
         * For details see http://download.oracle.com/javase/tutorial/uiswing/lookandfeel/plaf.html 
         */
        try {
            for (javax.swing.UIManager.LookAndFeelInfo info : javax.swing.UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    javax.swing.UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (ClassNotFoundException ex) {
            java.util.logging.Logger.getLogger(impor_ex.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (InstantiationException ex) {
            java.util.logging.Logger.getLogger(impor_ex.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (IllegalAccessException ex) {
            java.util.logging.Logger.getLogger(impor_ex.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        } catch (javax.swing.UnsupportedLookAndFeelException ex) {
            java.util.logging.Logger.getLogger(impor_ex.class.getName()).log(java.util.logging.Level.SEVERE, null, ex);
        }
        //</editor-fold>

        /* Create and display the form */
        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                try {
                    UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
                } catch (ClassNotFoundException ex) {
                    Logger.getLogger(impor_ex.class.getName()).log(Level.SEVERE, null, ex);
                } catch (InstantiationException ex) {
                    Logger.getLogger(impor_ex.class.getName()).log(Level.SEVERE, null, ex);
                } catch (IllegalAccessException ex) {
                    Logger.getLogger(impor_ex.class.getName()).log(Level.SEVERE, null, ex);
                } catch (UnsupportedLookAndFeelException ex) {
                    Logger.getLogger(impor_ex.class.getName()).log(Level.SEVERE, null, ex);
                }
                new impor_ex().setVisible(true);
            }
        });
        
        
    }

    // Variables declaration - do not modify//GEN-BEGIN:variables
    private javax.swing.JButton btn_Importar;
    private com.toedter.calendar.JDateChooser calendario;
    private javax.swing.JRadioButton deseleccion;
    private javax.swing.JComboBox<String> grupo;
    private javax.swing.JButton imprimir;
    private javax.swing.JLabel jLabel1;
    private javax.swing.JLabel jLabel2;
    private javax.swing.JLabel jLabel3;
    private javax.swing.JLabel jLabel4;
    private javax.swing.JScrollPane jScrollPane1;
    private javax.swing.ButtonGroup recuadros_Seleccion;
    private javax.swing.JButton salir;
    private javax.swing.JRadioButton seleccion;
    private javax.swing.JComboBox<String> semestre;
    private javax.swing.JTable tabla;
    // End of variables declaration//GEN-END:variables
}
