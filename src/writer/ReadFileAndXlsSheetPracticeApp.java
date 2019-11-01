package writer;



import java.io.*;

public class ReadFileAndXlsSheetPracticeApp {
     public static void main( String[] args ) throws IOException, FileNotFoundException
     {
         String fileName =  "testOutput.txt";
         try  {
             FileInputStream inputAFile = new FileInputStream(new File(fileName));
             System.out.println("filename = "+fileName);

             File forFileInputFile = new File(fileName);

            FileReader fileReader = new FileReader(fileName);
             System.out.println("fileReader ="+fileName);
         }catch (Exception e){
             System.out.println("in catch block");
         }



    }//end main method
}//end class
























