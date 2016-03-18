/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package scouting;

import java.io.File;
import java.io.IOException;
import java.util.logging.Level;
import java.util.logging.Logger;
import jxl.*;
import jxl.write.*;
import jxl.write.biff.RowsExceededException; 
import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.scene.layout.Pane;
import javafx.scene.paint.Color;
import javafx.stage.Stage;
import jxl.read.biff.BiffException;

/**
 * 
 * Cadillac Connectors team 5086 Scouting Program
 * 
 * Compiles data on other teams such as goals and defenses crossed using the Jexcel api.
 * 
 * Requires a file in the same directory entitled "final.xls", with the num "0" in every column for the first 40 or so
 * 
 * Created on 3/6/16
 * 
 * Last modified on 3/16/16
 *
 * @author Josh
 */
public class Scouting extends Application {
    
    public TextField teamNumber;
    public TextField matchNumber;
    public TextField dataA;
    public TextField dataB;
    public TextField dataC;
    public TextField dataD;
    public TextField dataE;
    public TextField dataF;
    public TextField dataG;
    public TextField dataH;
    public TextField dataI;
    public TextField dataJ;
    public TextField dataK;
    public TextField dataL;
    public TextField dataM;
    
    @Override
    //The forum locating all text fields and buttons
    public void start(Stage primaryStage) {
        
        //Data everything we want to gather on a team
        teamNumber = new TextField();
        teamNumber.setPromptText("Team Number");
        matchNumber = new TextField();
        matchNumber.setPromptText("Match Number");
        dataA = new TextField();
        dataA.setPromptText("High Goals Made");
        dataB = new TextField();
        dataB.setPromptText("High Goals Shot");
        dataC = new TextField();
        dataC.setPromptText("Low Goals Made");
        dataD = new TextField();
        dataD.setPromptText("Low Goals Shot");
        dataE = new TextField();
        dataE.setPromptText("Low Bars");
        dataF = new TextField();
        dataF.setPromptText("Portcullis");
        dataG = new TextField();
        dataG.setPromptText("Chivel de Frise");
        dataH = new TextField();
        dataH.setPromptText("Ramparts");
        dataI = new TextField();
        dataI.setPromptText("Moats");
        dataJ = new TextField();
        dataJ.setPromptText("Drawbridges");
        dataK = new TextField();
        dataK.setPromptText("Sally Ports");
        dataL = new TextField();
        dataL.setPromptText("Rock Walls");
        dataM = new TextField();
        dataM.setPromptText("Rough Terrains");
        TextField checkField =new TextField();
        TextField link = new TextField();
        link.setText("data:image/png;base64,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");
        
        //Submit to private sheet
        Button send = new Button();
        send.setText("Send to the sheet");
        
        //Compile in a sheet called "final.xls"
        Button finish = new Button();
        finish.setText("Finalize this Team");
        
        //Check to see how much data we have on the team
        Button check = new Button();
        check.setText("Check for data");
        
        //Submits data from the fields to a spreadsheet
        send.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                
                //Compile the data
                Scouting sc = new Scouting();
                String num = teamNumber.getText();
                String match = matchNumber.getText();
                String data1 = dataA.getText();
                String data2 = dataB.getText();
                String data3 = dataC.getText();
                String data4 = dataD.getText();
                String data5 = dataE.getText();
                String data6 = dataF.getText();
                String data7 = dataG.getText();
                String data8 = dataH.getText();
                String data9 = dataI.getText();
                String data10= dataJ.getText();
                String data11= dataK.getText();
                String data12= dataL.getText();
                String data13= dataM.getText();
                
                //Atempts to Send the data to the sheet by calling the sendToSheet() method
                //I know the error collection is only default and not great
                try {
                    sc.sendToSheet(num, match, data1, data2, data3, data4, data5, data6, data7, data8, data9, data10, data11, data12, data13);
                } catch (WriteException | IOException | BiffException ex) {
                    Logger.getLogger(Scouting.class.getName()).log(Level.SEVERE, null, ex);
                }
                
                teamNumber.clear();
                matchNumber.clear();
                dataA.clear();
                dataB.clear();
                dataC.clear();
                dataD.clear();
                dataE.clear();
                dataF.clear();
                dataG.clear();
                dataH.clear();
                dataI.clear();
                dataJ.clear();
                dataK.clear();
                dataL.clear();
                dataM.clear();
                
            }
        });
        
        //Sends the current team number to the main spreadsheet
        finish.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                //Get the team number
                Scouting sc = new Scouting();
                String num = teamNumber.getText();
                
                //Try to compile it using the compileSheet() method
                //I know the error collection is only default and not great
                try {
                    sc.compileSheet(num);
                } catch (WriteException | IOException | BiffException ex) {
                    Logger.getLogger(Scouting.class.getName()).log(Level.SEVERE, null, ex);
                }
                
            }
        });
        
        //Checks how much data has been imputed for the given team
        check.setOnAction(new EventHandler<ActionEvent>() {
            @Override
            public void handle(ActionEvent event) {
                int used=0;
                try {
                    Workbook get = Workbook.getWorkbook(new File(teamNumber.getText() + ".xls"));
                    if (!"".equals(get.getSheet(0).getCell(0, 0).getContents())) {
                        used=1;
                    }
                    if (!"".equals(get.getSheet(0).getCell(1, 0).getContents())) {
                        used=2;
                    }
                    if (!"".equals(get.getSheet(0).getCell(2, 0).getContents())) {
                        used=3;
                    }
                    if (!"".equals(get.getSheet(0).getCell(3, 0).getContents())) {
                        used=4;
                    }
                    if (!"".equals(get.getSheet(0).getCell(4, 0).getContents())) {
                        used=5;
                    }
                    if (!"".equals(get.getSheet(0).getCell(5, 0).getContents())) {
                        used=6;
                    }
                    if (!"".equals(get.getSheet(0).getCell(6, 0).getContents())) {
                        used=7;
                    }
                    if (!"".equals(get.getSheet(0).getCell(7, 0).getContents())) {
                        used=8;
                    }
                    if (!"".equals(get.getSheet(0).getCell(8, 0).getContents())) {
                        used=9;
                    }
                    if (!"".equals(get.getSheet(0).getCell(9, 0).getContents())) {
                        used=10;
                    }
                    if (!"".equals(get.getSheet(0).getCell(10, 0).getContents())) {
                        used=11;
                    }
                    if (!"".equals(get.getSheet(0).getCell(11, 0).getContents())) {
                        used=12;
                    }
                    if (!"".equals(get.getSheet(0).getCell(12, 0).getContents())) {
                        used=12;
                    }
                    
                String output = Integer.toString(used);
                checkField.setText(output);
                } catch (IOException | BiffException ex) {
                    teamNumber.clear();
                    checkField.setText("ERROR 404 - Team Not Found");
                }
            }
        });
        
        //The default forum
        Pane root = new Pane();
        
        //Sets the location and size of the buttons
        send.setLayoutX(0);
        send.setLayoutY(375);
        send.setMinSize(150, 25);
        root.getChildren().add(send);
        finish.setLayoutX(0);
        finish.setLayoutY(400);
        finish.setMinSize(150, 25);
        root.getChildren().add(finish);
        check.setLayoutX(150);
        check.setLayoutY(0);
        check.setMinSize(150, 25);
        
        //Sets the location of the metric butt-ton of Text Fields
        teamNumber.setLayoutX(0);
        teamNumber.setLayoutY(0);
        root.getChildren().add(teamNumber);
        root.getChildren().add(check);
        checkField.setLayoutX(150);
        checkField.setLayoutY(25);
        root.getChildren().add(checkField);
        matchNumber.setLayoutX(0);
        matchNumber.setLayoutY(25);
        root.getChildren().add(matchNumber);
        dataA.setLayoutX(0);
        dataA.setLayoutY(50);
        root.getChildren().add(dataA);
        dataB.setLayoutX(0);
        dataB.setLayoutY(75);
        root.getChildren().add(dataB);
        dataC.setLayoutX(0);
        dataC.setLayoutY(100);
        root.getChildren().add(dataC);
        dataD.setLayoutX(0);
        dataD.setLayoutY(125);
        root.getChildren().add(dataD);
        dataE.setLayoutX(0);
        dataE.setLayoutY(150);
        root.getChildren().add(dataE);
        dataF.setLayoutX(0);
        dataF.setLayoutY(175);
        root.getChildren().add(dataF);
        dataG.setLayoutX(0);
        dataG.setLayoutY(200);
        root.getChildren().add(dataG);
        dataH.setLayoutX(0);
        dataH.setLayoutY(225);
        root.getChildren().add(dataH);
        dataI.setLayoutX(0);
        dataI.setLayoutY(250);
        root.getChildren().add(dataI);
        dataJ.setLayoutX(0);
        dataJ.setLayoutY(275);
        root.getChildren().add(dataJ);
        dataK.setLayoutX(0);
        dataK.setLayoutY(300);
        root.getChildren().add(dataK);
        dataL.setLayoutX(0);
        dataL.setLayoutY(325);
        root.getChildren().add(dataL);
        dataM.setLayoutX(0);
        dataM.setLayoutY(350);
        root.getChildren().add(dataM);
        link.setLayoutX(150);
        link.setLayoutY(50);
        root.getChildren().add(link);
        
        //Sets the size of the forum & other little details
        Scene scene = new Scene(root, 300, 425);
        primaryStage.setTitle("Robot Scout Program");
        primaryStage.setScene(scene);
        primaryStage.show();
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        launch(args);
    }
    
    /**
     * 
     * Compiles data and imputes it into a spreadsheet.
     * I know that each of the little subtasks should be it's own method, but I'm lazy.
     * 
     * @param teamNum The team Number
     * @param matchNum The match Number
     * @param dataA Some data we want to take from the teams
     * @param dataB Some data we want to take from the teams
     * @param dataC Some data we want to take from the teams
     * @param dataD Some data we want to take from the teams
     * @param dataE Some data we want to take from the teams
     * @param dataF Some data we want to take from the teams
     * @param dataG Some data we want to take from the teams
     * @param dataH Some data we want to take from the teams
     * @param dataI Some data we want to take from the teams
     * @param dataJ Some data we want to take from the teams
     * @param dataK Some data we want to take from the teams
     * @param dataL Some data we want to take from the teams
     * @param dataM Some data we want to take from the teams
     * 
     * @author Josh
     * 
    **/
    public void sendToSheet (
            String teamNum,
            String matchNum,
            String dataA,
            String dataB,
            String dataC,
            String dataD,
            String dataE,
            String dataF,
            String dataG,
            String dataH,
            String dataI,
            String dataJ,
            String dataK,
            String dataL,
            String dataM
        )   throws RowsExceededException, WriteException, IOException, BiffException {
        
        
        File f = new File(teamNum + ".xls");
        File fb = new File(teamNum + "b.xls");
        
        //Copies the old file of the specified team num if it exists, and inputs new data
        if(f.exists() && !f.isDirectory()) {
        Workbook workbook; 
        workbook = Workbook.getWorkbook(f);
        
        WritableWorkbook newBook;
        newBook = Workbook.createWorkbook(fb);
        WritableSheet sheet = newBook.createSheet(teamNum, 0);
        Sheet old = workbook.getSheet(0);
        int indexValueX = 0;
        int indexValueY;
        Label N;
        String movingNumber;
        Cell oldCell;
        int indexValueC;
        
        //Copies old sheet over 1 column to make room for the new data
        while (indexValueX<=15) {
        indexValueY=15;
        
        while (indexValueY > 0) {
            indexValueC = indexValueY-1;
            oldCell = old.getCell(indexValueC, indexValueX);
            movingNumber = oldCell.getContents();
            N = new Label(indexValueY,indexValueX,movingNumber);
            sheet.addCell(new Label(N));
            indexValueY=indexValueY-1;
        }
        
        indexValueX=indexValueX+1;
        }
        
        //Gets the new data and puts it into the sheet
        Label dataLabel = new Label(0, 0, teamNum);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 1, matchNum);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 2, dataA);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 3, dataB);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 4, dataC);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 5, dataD);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 6, dataE);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 7, dataF);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 8, dataG);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 9, dataH);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,10, dataI);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,11, dataJ);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,12, dataK);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,13, dataL);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,14, dataM);
        sheet.addCell(dataLabel);
        
        
        //Finishes the books
        newBook.write();
        newBook.close();
        workbook.close();
        
        //Deletes the old and replaces it with the new
        f.delete();
        fb.renameTo(f);
        
        }
        //If there is no sheet with the team number name, this will create it
        else {
            
        //Creates a new sheet with the name of the team number
        WritableWorkbook newBook = Workbook.createWorkbook(new File (teamNum + ".xls"));
        WritableSheet sheet = newBook.createSheet(teamNum, 0);
        
        //Adds the new data to the newly created sheet
        Label dataLabel = new Label(0, 0, teamNum);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 1, matchNum);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 2, dataA);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 3, dataB);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 4, dataC);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 5, dataD);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 6, dataE);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 7, dataF);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 8, dataG);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0, 9, dataH);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,10, dataI);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,11, dataJ);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,12, dataK);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,13, dataL);
        sheet.addCell(dataLabel);
        dataLabel = new Label(0,14, dataM);
        sheet.addCell(dataLabel);
        
        //To quote Shia LeBouf, "Just do it!"
        sheet.addCell(new Label(25,25,"Give me candy"));
        
        //Finishes the new book
        newBook.write();
        newBook.close();
        
        }
        
        //Resets the boxes
        
        
    }
    
    /**
     * Compiles all the data from a team and enters it into the main sheet.
     * 
     * @param teamNum Team number to compile
     */
    public void compileSheet(String teamNum) throws IOException, BiffException, WriteException {
        Workbook oldBook = Workbook.getWorkbook(new File(teamNum + ".xls"));
        Sheet oldSheet = oldBook.getSheet(0);
        File f = new File("final.xls");
        File f2 = new File("finalb.xls");
        Workbook finalBookOld = Workbook.getWorkbook(f);
        Sheet finalSheetOld = finalBookOld.getSheet(0);
        WritableWorkbook finalBook = Workbook.createWorkbook(f2);
        WritableSheet finalSheet= finalBook.createSheet("Final Data", 0);
        int whatRowDoThePeopleUse = 1;
        int actualRowToUse = 1000;
        int yIndex = 2;
        double lowMade = 1;
        double lowShot = 1;
        double highMade= 1;
        double highShot= 1;
        int numrows = finalSheetOld.getRows();
        int numcols = finalSheetOld.getColumns();
        int copyY=0;
        int copyX=0;
        double lowPercent;
        double highPercent;
        
        //Copies the old Sheet
        while (copyX<numrows) {
            copyY=0;
            while (copyY<numcols) {
                Cell oldCell = finalSheetOld.getCell(copyY, copyX);
                String oldContents = oldCell.getContents();
                finalSheet.addCell(new Label(copyY, copyX, oldContents));
                copyY++;
            }
            copyX++;
        }
        
        //Finds out what column to imput the data into
        while (whatRowDoThePeopleUse < 1000) {
            Cell old = finalSheet.getCell(whatRowDoThePeopleUse, 0);
            String isItZero = old.getContents();
            if ("0".equals(isItZero)) {
                actualRowToUse=whatRowDoThePeopleUse;
                whatRowDoThePeopleUse=1001;
            }
            whatRowDoThePeopleUse=whatRowDoThePeopleUse+1;
        }
        
        //adds data from all the original cells, starting with cell 3
        while (yIndex < 15) {
        Cell cell1 = oldSheet.getCell(0, yIndex);
        String string1 = cell1.getContents();
        Cell cell2 = oldSheet.getCell(1, yIndex);
        String string2 = cell2.getContents();
        Cell cell3 = oldSheet.getCell (2, yIndex);
        String string3 = cell3.getContents();
        Cell cell4 = oldSheet.getCell (3, yIndex);
        String string4 = cell4.getContents();
        Cell cell5 = oldSheet.getCell (4, yIndex);
        String string5 = cell5.getContents();
        Cell cell6 = oldSheet.getCell (5, yIndex);
        String string6 = cell6.getContents();
        Cell cell7 = oldSheet.getCell (6, yIndex);
        String string7 = cell7.getContents();
        Cell cell8 = oldSheet.getCell (7, yIndex);
        String string8 = cell8.getContents();
        Cell cell9 = oldSheet.getCell (8, yIndex);
        String string9 = cell9.getContents();
        Cell cell10 = oldSheet.getCell (9, yIndex);
        String string10 = cell10.getContents();
        Cell cell11 = oldSheet.getCell (10, yIndex);
        String string11 = cell11.getContents();
        Cell cell12 = oldSheet.getCell (11, yIndex);
        String string12 = cell12.getContents();
        
        //Converts n cells for cannot cross to -1
        if ("n".equals(string1)) {
            string1="-1";
        }
        if ("n".equals(string2)) {
            string2="-1";
        }
        if ("n".equals(string3)) {
            string3="-1";
        }
        if ("n".equals(string4)) {
            string4="-1";
        }
        if ("n".equals(string5)) {
            string5="-1";
        }
        if ("n".equals(string6)) {
            string6="-1";
        }    
        if ("n".equals(string7)) {
            string7="-1";
        }
        if ("n".equals(string8)) {
            string8="-1";
        }
        if ("n".equals(string9)) {
            string9="-1";
        }
        if ("n".equals(string10)) {
            string10="-1";
        }
        if ("n".equals(string11)) {
            string11="-1";
        }
        if ("n".equals(string12)) {
            string12="-1";
        }
        
        //Converts . cells for cannot cross to -1
        if (".".equals(string1)) {
            string1="-1";
        }
        if (".".equals(string2)) {
            string2="-1";
        }
        if (".".equals(string3)) {
            string3="-1";
        }
        if (".".equals(string4)) {
            string4="-1";
        }
        if (".".equals(string5)) {
            string5="-1";
        }
        if (".".equals(string6)) {
            string6="-1";
        }    
        if (".".equals(string7)) {
            string7="-1";
        }
        if (".".equals(string8)) {
            string8="-1";
        }
        if (".".equals(string9)) {
            string9="-1";
        }
        if (".".equals(string10)) {
            string10="-1";
        }
        if (".".equals(string11)) {
            string11="-1";
        }
        if (".".equals(string12)) {
            string12="-1";
        }
        
        //Converts blank cells to 0 because Irwin is lazy
        if ("".equals(string1)) {
            string1="0";
        }
        if ("".equals(string2)) {
            string2="0";
        }
        if ("".equals(string3)) {
            string3="0";
        }
        if ("".equals(string4)) {
            string4="0";
        }
        if ("".equals(string5)) {
            string5="0";
        }
        if ("".equals(string6)) {
            string6="0";
        }    
        if ("".equals(string7)) {
            string7="0";
        }
        if ("".equals(string8)) {
            string8="0";
        }
        if ("".equals(string9)) {
            string9="0";
        }
        if ("".equals(string10)) {
            string10="0";
        }
        if ("".equals(string11)) {
            string11="0";
        }
        if ("".equals(string12)) {
            string12="0";
        }
        
        //Converts the strings to ints to add them
        int int1 = Integer.valueOf(string1);
        int int2 = Integer.valueOf(string2);
        int int3 = Integer.valueOf(string3);
        int int4 = Integer.valueOf(string4);
        int int5 = Integer.valueOf(string5);
        int int6 = Integer.valueOf(string6);
        int int7 = Integer.valueOf(string7);
        int int8 = Integer.valueOf(string8);
        int int9 = Integer.valueOf(string9);
        int int10 = Integer.valueOf(string10);
        int int11 = Integer.valueOf(string11);
        int int12 = Integer.valueOf(string12);
        
        //Adds the values
        int finalCount = int1+int2+int3+int4+int5+int6+int7+int8+int9+int10+int11+int12;
        
        //Attempts to add that value to the new column in "final.xls"
            try {
                finalSheet.addCell(new jxl.write.Number(actualRowToUse, yIndex, finalCount));
            } catch (WriteException ex) {
                Logger.getLogger(Scouting.class.getName()).log(Level.SEVERE, null, ex);
            }
            
        //Stores values from the high and low goals
            if (yIndex==2) {
                lowMade = finalCount;
            }
            if (yIndex==3) {
                lowShot = finalCount;
            }
            if (yIndex==4) {
                highMade = finalCount;
            }
            if (yIndex==5) {
                highShot = finalCount;
            }
            
            //Increment the yIndex to use the next row
            yIndex=yIndex+1;
        }
        
        
        //Avoids Divide by zero error
        if (lowShot == 0) {
            lowPercent=0;
        }
        //Gets the shot percentage for low goals
        else {
            lowPercent = (lowMade/lowShot);
        }
        
        //Avoids Divide by zero error
        if (highShot == 0) {
            highPercent=0;
        }
        //Gets the shot percentage for high goals
        else {
            highPercent = (highMade/highShot);
        }
        
        //Converts the doubles to strings. I have no clue why I did this
        String highStrung = Double.toString(highPercent);
        String lowStrung  = Double.toString(lowPercent);
        
        //Sends the percents to the sheet
        finalSheet.addCell(new Label(actualRowToUse, 3, lowStrung));
        finalSheet.addCell(new Label(actualRowToUse, 5, highStrung));
        
        //Sends the team number and adds a placeholder 1 to the sheet
        finalSheet.addCell(new Label(actualRowToUse, 1, teamNum));
        finalSheet.addCell(new Label(actualRowToUse, 0, "1"));
        
        //Finalizes the book
        finalBook.write();
        finalBook.close();
        
        
        //Deletes the old "final.xls"
        f.delete();
        
        //Renames "finalb.xls" to "final.xls"
        f2.renameTo(f);
        
    }
}//Josh likes cookies