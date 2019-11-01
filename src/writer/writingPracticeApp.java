package writer;

import java.io.FileNotFoundException;
import java.io.PrintWriter;

public class writingPracticeApp {



    public static void main( String[] args ) {



         String namedFile = "readerFile1.txt";
        try {
            PrintWriter outStreamFile = new PrintWriter(namedFile);
            outStreamFile.println("You've created and written to File: \n"+ namedFile);//stores in ram until pushed to file or closed...
            outStreamFile.flush();//forces data to the file.
            outStreamFile.println("more test written");
            outStreamFile.close();// .close() to close or .flush() to keep writing and ensure push to file...
            System.out.println("printed: "+namedFile);

        } catch (FileNotFoundException e) {
            System.out.println("throws file not found exception.. needs valid file");
            e.printStackTrace();
        }


    }






}
