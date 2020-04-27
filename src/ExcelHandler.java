/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author User
 */
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashSet;
import java.util.Locale;
import javax.swing.table.DefaultTableModel;
import jxl.Cell;
import jxl.CellType;
import jxl.write.Number;
import java.util.Scanner;
import jxl.Sheet;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.read.biff.BiffException;
import jxl.write.DateFormats;
import jxl.write.DateTime;
import jxl.write.Formula;
import jxl.write.Label;
import jxl.write.NumberFormat;
import jxl.write.NumberFormats;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import org.openide.util.Exceptions;

public class ExcelHandler
{

    private String inputFile;

    public ExcelHandler(String inputFile)
    {
        this.inputFile = inputFile;
    }

    public void setInputFile(String inputFile)
    {
        this.inputFile = inputFile;
    }
    
    /*
     * Assumes Data appears in 0th Sheet from row 1 (row 0 is header). column serial is sourceID, targetID, sourceIsAborg, targetIsAborg and rest columns mean different network data as indicated by totalColumnNumber. 
     * Last totalColumnToConsiderAsFreq will indicate how many columns from the beginning of freq columns as having values apart from 0 and 1
     */
    public void prepareMargaretDataset(String fileNameToRead, String fileNameToWrite, int totalColumnNumber, int totalColumnToConsiderAsFreq)
    {
        ArrayList<String> nodeInventory = new ArrayList<>();
        ArrayList<String> fromNodeID = new ArrayList<>();
        ArrayList<String> toNodeID = new ArrayList<>();
        ArrayList<Integer> isAboriginal = new ArrayList<>();
        ArrayList<Integer> isRespondent = new ArrayList<>();
        ArrayList<Integer> inDegree = new ArrayList<>();
        
        try
        {
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            WritableCellFormat cf = new WritableCellFormat(wf);
            cf.setWrap(true);  
            WritableSheet s,sInDegree;
            File inputWorkbook = new File(fileNameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            Sheet readSheet = wbToRead.getSheet(0) ;//read the first sheet
            
            //read the source and target nodeID's and their aboriginality and build the nodeList
            String tempSource, tempTarget;
            int tempAborgSource, tempAborgTarget, tempTieStrength;
            for(int j=1;j<readSheet.getRows();j++)
            {
                tempSource = readSheet.getCell(0,j).getContents().trim();
                tempTarget = readSheet.getCell(1,j).getContents().trim();
                tempAborgSource = Integer.parseInt(readSheet.getCell(2,j).getContents());
                tempAborgTarget = Integer.parseInt(readSheet.getCell(3,j).getContents());
                if(!nodeInventory.contains(tempSource))//new Node, insert
                {
                    nodeInventory.add(tempSource);
                    isAboriginal.add(tempAborgSource);
                    isRespondent.add(1);
                }
                else
                {//this source is respondent. So update it if not already done so.
                    int i = nodeInventory.indexOf(tempSource);
                    isRespondent.set(i, 1);
                }
                if(!tempSource.equalsIgnoreCase(tempTarget) && !nodeInventory.contains(tempTarget))
                {
                    nodeInventory.add(tempTarget);
                    isAboriginal.add(tempAborgTarget);
                    isRespondent.add(0);
                }
            }//finished populating inventory, 

            s = workbook.createSheet("Node Inventory", 0);
            s.addCell(new Label(0,0,"Node ID"));
            s.addCell(new Label(1,0,"Aboriginality"));
            s.addCell(new Label(2,0,"Respondent"));
            
            for(int i =0;i<nodeInventory.size();i++)
            {
                s.addCell(new Label(0,i+1,nodeInventory.get(i)));    
                s.addCell(new Number(1,i+1,isAboriginal.get(i)));    
                s.addCell(new Number(2,i+1,isRespondent.get(i)));    
            }//finish writing 1st sheet, node Inventory. it'll later be written to include inDegree columns
            sInDegree = workbook.getSheet(0);//sIndegree will write inDegree columns in iterations
            
            //loop through all other networks and write in different sheets;
            for(int i=0;i<totalColumnNumber;i++)
            {
                inDegree.clear();//clear the list                
                for(int j=0;j<nodeInventory.size();j++)
                    inDegree.add(0);
                
                s = workbook.createSheet(readSheet.getCell(4+i,0).getContents(), 1+i);//giving the new sheet name as the column header name
                s.addCell(new Label(0,0,"From"));
                s.addCell(new Label(1,0,"To"));
                s.addCell(new Label(2,0,"Strength"));
                for(int j=1,k=1;j<readSheet.getRows();j++)
                {
                    tempSource = readSheet.getCell(0, j).getContents();
                    tempTarget = readSheet.getCell(1,j).getContents();
                    try{
                    tempTieStrength = Integer.parseInt(readSheet.getCell(4+i,j).getContents());
                    }
                    catch(NumberFormatException nf)
                    {
                        System.out.println("Probably null at col row " + (4+i) + " " + j);
                        tempTieStrength = 0;
                    }
                    //include only if tieStrength is non-zero
                    if(tempTieStrength > 0)
                    {
                        s.addCell(new Label(0, k, tempSource));
                        s.addCell(new Label(1,k, tempTarget));
                        if(i < totalColumnToConsiderAsFreq)
                        {
                            s.addCell(new Number(2,k,tempTieStrength));
                            inDegree.set(nodeInventory.indexOf(tempTarget), inDegree.get(nodeInventory.indexOf(tempTarget))+tempTieStrength);
                        
                        }
                        else
                        {
                            s.addCell(new Number(2,k,1));
                            inDegree.set(nodeInventory.indexOf(tempTarget), inDegree.get(nodeInventory.indexOf(tempTarget))+1);
                        }
                        k++;
                    }
                }//completed doing one sheet. Write inDegree column and then Do another
                sInDegree.addCell(new Label(4+i,0,readSheet.getCell(4+i,0).getContents()));
                for(int j=0;j<nodeInventory.size();j++)
                    sInDegree.addCell(new Number(4+i,j+1,inDegree.get(j)));
            }
            
            
            workbook.write();
            workbook.close(); 
            System.out.println("Done writing");
        }
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }
    public void calculateANForEmailDataset(String filenameToRead, String fileNameToWrite)
    {
        int totalInputSheet = 6;
        int from, to;
        double freq;
        int firstOccurrence, lastOccurrence;
        ArrayList <Integer> fromNodeID = new ArrayList<>();//   1 1 1 3 2 2
        ArrayList <Integer> toNodeID = new ArrayList<>();//     2 3 4 4 1 2
        ArrayList <Double> freqOfOccurrence = new ArrayList<>();
        try
        {
          File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            Sheet readSheet ;  
            for(int i=0;i<totalInputSheet;i++)
            {
                readSheet = wbToRead.getSheet(i+1);
                for (int j = 1; j < readSheet.getRows(); j++)
                {
                    from = Integer.parseInt(readSheet.getCell(0, j).getContents());
                    to = Integer.parseInt(readSheet.getCell(1, j).getContents());
                    freq = Double.parseDouble(readSheet.getCell(2, j).getContents()); 
                    //search if from, to edge is found in two Arraylist, this means the edge is present
                    firstOccurrence = fromNodeID.indexOf(from);//should get first occurrence ID
                    if(firstOccurrence != -1)//from ID is present, search all of from ID's
                    {
                        lastOccurrence = fromNodeID.lastIndexOf(from);//now search from firstOccurrence to last for the other part of the edge
                        int k;
                        for(k=firstOccurrence;k<=lastOccurrence;k++)
                        {
                            if(toNodeID.get(k)==to)//got a match, the edge is present
                            {
                                freqOfOccurrence.set(k, freqOfOccurrence.get(k)+freq);//added the freq
                                break;
                            }
                        }
                        if(k>lastOccurrence)//we couldn't find an edge, so insert it after lastOccurrence
                        {
                            fromNodeID.add(lastOccurrence+1,from);
                            toNodeID.add(lastOccurrence+1,to);
                            freqOfOccurrence.add(lastOccurrence+1,freq);
                        }
                    }
                    else//the edge is not present. insert it at the last
                    {
                        fromNodeID.add(from);
                        toNodeID.add(to);
                        freqOfOccurrence.add(freq);
                    }
                }//end traversing rows of current sheet
                System.out.println("Finished processing month " + i);
            }//end traversing all the required sheets
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            WritableCellFormat cf = new WritableCellFormat(wf);
            WritableCellFormat wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            cf.setWrap(true);
            WritableSheet s = workbook.createSheet("Index", 0);
            
            for(int i =0;i<fromNodeID.size();i++)
            {
                s.addCell(new Number(0,i,fromNodeID.get(i)));    
                s.addCell(new Number(1,i,toNodeID.get(i)));    
                s.addCell(new Number(2,i,freqOfOccurrence.get(i)));    
            }
            workbook.write();
            workbook.close();
            System.out.println("DONE");
        }//end trying read and/or write
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }
    public void calculateTotalIncomingOutgoingforEmailDataset(String filenameToRead, String fileNameToWrite)
    {
        int totalInputSheet = 26;
        int from, to, freq;
        ArrayList <Integer> freqListOutgoing = new ArrayList<>();
        ArrayList <Integer> freqListIncoming = new ArrayList<>();
        for(int i = 0;i<6511;i++)
        {
            freqListOutgoing.add(0);
            freqListIncoming.add(0);
        }
        try
        {
            File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            Sheet readSheet ;
            for(int i=0;i<totalInputSheet;i++)//
            {
                readSheet = wbToRead.getSheet(i+1);
                //s = workbook.createSheet("Week" + (i+27), (i+1));
                //s.addCell(new Label(0,0,"FROM ID",cf));
                //s.addCell(new Label(1,0,"TO ID",cf));
                //s.addCell(new Label(2,0,"FREQ",cf));
                
                for (int j = 1; j < readSheet.getRows(); j++)
                {
                    from = Integer.parseInt(readSheet.getCell(0, j).getContents());
                    to = Integer.parseInt(readSheet.getCell(1, j).getContents());
                    freq = Integer.parseInt(readSheet.getCell(2, j).getContents());  
                    freqListOutgoing.set(from, freqListOutgoing.get(from) + freq);
                    freqListIncoming.set(to, freqListIncoming.get(to) + freq);
                }
                System.out.println("Finished processing week " + (27+i));
            }
            
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            WritableCellFormat cf = new WritableCellFormat(wf);
            WritableCellFormat wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            cf.setWrap(true);
            WritableSheet s = workbook.createSheet("Index", 0);
            
            for(int i =0;i<freqListOutgoing.size();i++)
            {
                s.addCell(new Number(0,i,freqListOutgoing.get(i)));    
                s.addCell(new Number(1,i,freqListIncoming.get(i)));    
            }
            workbook.write();
            workbook.close();
            System.out.println("DONE");
        }
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }
    public void compensateDegreeForSINs(String filenameToRead, String fileNameToWrite)
    {
        int totalInputSheet = 26;//week 27 to 52, total 26
        int totalActor = 6511;//0 to 6510
        WritableSheet s;
        Sheet readSheet;
        int from, to, freq;
        ArrayList<Integer> SINDegree = new ArrayList<>();
        try
        {
            File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            s = workbook.createSheet("Merged", 0);
            
            for(int i = 0;i< totalInputSheet;i++)//loop through ALL SIN. 
            {   //reset the ArrayList
                SINDegree.clear();
                for(int j=0;j<totalActor;j++)
                {
                    SINDegree.add(0);//later 0 will be inserted as -1
                }
                readSheet = wbToRead.getSheet(i+2);  //SIN Starts from 2
                for(int j=1;j<readSheet.getRows();j++)//traverse all the rows, except the header row
                {
                    from = Integer.parseInt(readSheet.getCell(0,j).getContents());
                    to = Integer.parseInt(readSheet.getCell(1,j).getContents());
                    freq = Integer.parseInt(readSheet.getCell(2,j).getContents());
                    SINDegree.set(from,SINDegree.get(from)+ freq);
                    SINDegree.set(to,SINDegree.get(to)+ freq);
                }//finished one SIN calculation
                s.addCell(new Label(i, 0, "Degree Week " + (27+i)));
                for(int j=0;j<totalActor;j++)
                {
                    if(SINDegree.get(j)==0)
                        s.addCell(new Number(i,j+1,-1));
                    else
                        s.addCell(new Number(i,j+1,SINDegree.get(j)));
                }
                System.out.println("Finished week " + (27+i));
            }//finish looping all SIN
            workbook.write();
            workbook.close();   
            System.out.println("Finished readlly");
        }
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
        
    }
    public void doASCalculation(String filenameToRead, String locationToWrite)
    {
        try
        {
            String from, to;
            String Date = "Dummy", oldDate = "Dummy";
            Scanner reader = new Scanner(new File(filenameToRead));
            BufferedWriter SINout = null;
            while(reader.hasNext())
            {
                from = reader.next();
                to = (reader.next());
                Date = reader.next();
                if(!Date.equals(oldDate))
                {
                    //write new Text file
                    System.out.println("Writing new file " + Date);
                    oldDate = Date;
                    if(SINout != null)
                    {
                        //close the file
                        SINout.close();
                    }
                    SINout = new BufferedWriter(new FileWriter(locationToWrite+oldDate+".txt"));
                }
                SINout.newLine();
                SINout.write(from + "," + to);
            }
            SINout.close();
            System.out.println("Done writing. Hurray");
        } 
        catch (FileNotFoundException ex)
        {
            System.out.println("File Not found");
        }
        catch(IOException ex)
        {
            ex.printStackTrace();
        }
        catch(Exception ex)
        {
            ex.printStackTrace();
        }
    }
    public void copyAllSINResultInOneSheet(String filenameToRead, String fileNameToWrite)
    {
        int totalInputSheet = 6;//July to December
        int totalActor = 2253;//0 to 2252
        int currentActor;
        WritableSheet s;
        Sheet readSheet;
        int degCol, BtwCol, ClsCol;
        ArrayList<Double> SINDegree = new ArrayList<>();
        ArrayList<Double> SINBtwnns = new ArrayList<>();
        ArrayList<Double> SINClsns = new ArrayList<>();
        
        try
        {
            File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            s = workbook.createSheet("Merged", 0);
             
            
            for(int i=0;i<totalInputSheet;i++)//loop through ALL SIN
            {
                //reset the Arraylist
                SINDegree.clear();
                SINBtwnns.clear();
                SINClsns.clear();
                for(int j=0;j<totalActor;j++)
                {
                    SINDegree.add(-1.0);//-1 means emtpy
                    SINBtwnns.add(-1.0);
                    SINClsns.add(-1.0);
                }
                
                readSheet = wbToRead.getSheet(i+2);  //SIN Starts from 2
                for (int j=1;j<readSheet.getRows();j++)
                {
                    currentActor = Integer.parseInt(readSheet.getCell(0,j).getContents());
                    SINDegree.set(currentActor, Double.parseDouble(readSheet.getCell(2,j).getContents()));
                    SINBtwnns.set(currentActor, Double.parseDouble(readSheet.getCell(3,j).getContents()));
                    SINClsns.set(currentActor, Double.parseDouble(readSheet.getCell(4,j).getContents()));
                }
                degCol = i+0;
                BtwCol = i+7;
                ClsCol = i+14;
                s.addCell(new Label(degCol,0,"Degree Month " + i));
                s.addCell(new Label(BtwCol,0,"Btwnns Month " + i));
                s.addCell(new Label(ClsCol,0,"Clsns Month " + i));
                for(int j = 0; j<totalActor;j++)//traverse all SIN, If not found they starts from sheet 2
                {
                    s.addCell(new Number(degCol,j+1,SINDegree.get(j)));
                    s.addCell(new Number(BtwCol,j+1,SINBtwnns.get(j)));
                    s.addCell(new Number(ClsCol,j+1,SINClsns.get(j)));
                }
                System.out.println("end traversing Month " + i);
            }//end traversing all the sheets
            workbook.write();
            workbook.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }        
    }

    public void calculateDDAEmailDataset(String filenameToRead, String fileNameToWrite )
    {
        int totalInputSheet = 6;//month 0 to month 5, July to December
        WritableSheet s;
        double OVANDegree, OVANBtwnns, OVANClsns;
        double DDADegree,DDABtwnns,DDAClsns;
        int colDeg, colBtwn, colClsns;
        boolean prevPresentDegree, prevPresentBtwn, prevPresentCls;
        Sheet readSheet;
        try
        {
            File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            
            
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            WritableCellFormat cf = new WritableCellFormat(wf);
            WritableCellFormat wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            cf.setWrap(true);
            s = workbook.createSheet("DDA", 0);
            s.addCell(new Label(0,0,"ID"));
            s.addCell(new Label(1,0,"DDADegree"));
            s.addCell(new Label(2,0,"DDABtwnns"));
            s.addCell(new Label(3,0,"DDAClsns"));
                
            readSheet = wbToRead.getSheet(2);
            
            for(int i=1;i<readSheet.getRows();i++)//for all actors
            {
                OVANDegree = Double.parseDouble(readSheet.getCell(1, i).getContents());
                OVANBtwnns =  Double.parseDouble(readSheet.getCell(2, i).getContents());
                OVANClsns = Double.parseDouble(readSheet.getCell(3, i).getContents());
                DDADegree=DDABtwnns=DDAClsns = 0.0;//get ready
                
                prevPresentDegree = prevPresentBtwn = prevPresentCls = true;//initially think everyone is present
                for(int j = 0; j<totalInputSheet;j++)//traverse all SIN, they are all at the current sheet (2)
                {//for degree, SIN column 4+j, SIN BTWN 31+j, SIN clsns 58 + j
                    colDeg = 4+j;
                    colBtwn = 11+j;
                    colClsns = 18+j;
                    
                    if(Double.parseDouble(readSheet.getCell(colDeg,i).getContents()) > -1.0)//present, so consider
                    {
                        if(j==0)//first time
                            DDADegree = DDADegree + 1.0 * Math.abs(OVANDegree - Double.parseDouble(readSheet.getCell(colDeg,i).getContents()));
                        else
                            DDADegree = DDADegree + this.getAlpha(prevPresentDegree, true)* Math.abs(OVANDegree - Double.parseDouble(readSheet.getCell(colDeg,i).getContents()));                            
                        prevPresentDegree = true;
                    }
                    else
                    {
                        if(j == 0)//first time
                            DDADegree = DDADegree + 1.0 * Math.abs(OVANDegree);
                        else
                            DDADegree = DDADegree + this.getAlpha(prevPresentDegree, false)* Math.abs(OVANDegree);
                        prevPresentDegree = false;                        
                    }
                    if(Double.parseDouble(readSheet.getCell(colBtwn,i).getContents()) > -1.0)//present, so consider
                    {
                        if(j==0)//first time
                            DDABtwnns = DDABtwnns + 1.0 * Math.abs(OVANBtwnns - Double.parseDouble(readSheet.getCell(colBtwn,i).getContents()));
                        else
                            DDABtwnns = DDABtwnns + this.getAlpha(prevPresentBtwn, true)* Math.abs(OVANBtwnns - Double.parseDouble(readSheet.getCell(colBtwn,i).getContents()));
                        prevPresentBtwn = true;
                    }
                    else
                    {
                        if(j==0)
                            DDABtwnns = DDABtwnns + 1.0 * Math.abs(OVANBtwnns);
                        else
                            DDABtwnns = DDABtwnns + this.getAlpha(prevPresentBtwn, false)* Math.abs(OVANBtwnns);
                        prevPresentBtwn = false;
                    }
                    if(Double.parseDouble(readSheet.getCell(colClsns,i).getContents()) > -1.0)//present, so consider
                    {
                        if(j == 0)//for first time
                            DDAClsns = DDAClsns + 1.0 * Math.abs(OVANClsns - Double.parseDouble(readSheet.getCell(colClsns,i).getContents()));
                        else
                            DDAClsns = DDAClsns + this.getAlpha(prevPresentCls, true) * Math.abs(OVANClsns - Double.parseDouble(readSheet.getCell(colClsns,i).getContents()));
                        prevPresentCls = true;
                    }
                    else
                    {
                        if(j == 0)//for first time
                            DDAClsns = DDAClsns + 1.0 * Math.abs(OVANClsns);
                        else
                            DDAClsns = DDAClsns + this.getAlpha(prevPresentCls, false) * Math.abs(OVANClsns);
                        prevPresentCls = false;                        
                    }
                    
                }//end traversing the SIN's
                //finalize DDA by dividing by SIN number and write
                DDADegree = DDADegree / totalInputSheet;
                DDABtwnns = DDABtwnns / totalInputSheet;
                DDAClsns = DDAClsns / totalInputSheet;
                
                s.addCell(new Number(0,i,i-1));
                s.addCell(new Number(1,i,DDADegree));
                s.addCell(new Number(2,i,DDABtwnns));
                s.addCell(new Number(3,i,DDAClsns));
                System.out.println("Done Actor ID " + (i-1));
            }//end traversing one actor
            workbook.write();
            workbook.close();   
            System.out.println("It is Really finsihed");
        } catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }

    private double getAlpha(boolean previousPresent, boolean currentPresent)
    {     
        if(!currentPresent)//absent now, so 0
            return 0.0;
        else if(!previousPresent)//absent -> present
        {
            return 0.5;
        }
        else
            return 1.0;//present -> present
    }
    public void calculateDDAEmailDatasetNormalized(String filenameToRead, String fileNameToWrite )
    {
        int totalInputSheet = 6;//month 0 to month 5, July to December
        WritableSheet s;
        double OVANDegree;
        double DDADegree;
        int colDeg;
        boolean prevPresentDegree;
        Sheet readSheet;
        try
        {
            File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            
            
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            WritableCellFormat cf = new WritableCellFormat(wf);
            WritableCellFormat wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            cf.setWrap(true);
            s = workbook.createSheet("DDA", 0);
            s.addCell(new Label(0,0,"ID"));
            s.addCell(new Label(1,0,"DDADegree"));
                
            readSheet = wbToRead.getSheet(0);
            
            for(int i=1;i<2254;i++)//for all actors
            {
                OVANDegree = Double.parseDouble(readSheet.getCell(9, i).getContents());
             
                DDADegree= 0.0;//get ready
                
                prevPresentDegree = true;//initially think everyone is present
                for(int j = 0; j<totalInputSheet;j++)//traverse all SIN, they are all at the current sheet (2)
                {//for degree, SIN column 4+j, SIN BTWN 31+j, SIN clsns 58 + j
                    colDeg = 10+j;
                
                    
                    if(Double.parseDouble(readSheet.getCell(colDeg,i).getContents()) > -1.0)//present, so consider
                    {
                        if(j==0)//first time
                            DDADegree = DDADegree + 1.0 * Math.abs(OVANDegree - Double.parseDouble(readSheet.getCell(colDeg,i).getContents()));
                        else
                            DDADegree = DDADegree + this.getAlpha(prevPresentDegree, true)* Math.abs(OVANDegree - Double.parseDouble(readSheet.getCell(colDeg,i).getContents()));                            
                        prevPresentDegree = true;
                    }
                    else
                    {
                        if(j == 0)//first time
                            DDADegree = DDADegree + 1.0 * Math.abs(OVANDegree);
                        else
                            DDADegree = DDADegree + this.getAlpha(prevPresentDegree, false)* Math.abs(OVANDegree);
                        prevPresentDegree = false;                        
                    }
                }//end traversing the SIN's
                //finalize DDA by dividing by SIN number and write
                DDADegree = DDADegree / totalInputSheet;
                
                s.addCell(new Number(0,i,i-1));
                s.addCell(new Number(1,i,DDADegree));
                
                System.out.println("Done Actor ID " + (i-1));
            }//end traversing one actor
            workbook.write();
            workbook.close();   
            System.out.println("It is Really finsihed");
        } catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }
    
    /*
     * If the data is in one sheet and third column is time (31, 32 means week number, then isInOne = true and totalInputSheet = 1, time can't be minus
     * Input must be in column 1 (from), column 2(to), column 3(freq or time)
     */
    public void gCalculateIDAndFormatExcelStep1(int totalInputSheet, boolean isInOne, boolean removeDuplicate, String timeUnit, String filenameToRead, String fileNameToWrite)
    {
        ArrayList <String> nameList = new ArrayList<>();//for holding the names
        String from, to;
        WritableSheet s;
        int fromID, toID;
        int writingRow=1;
        double freq;
        
        try
        {
            File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            
            
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            s = workbook.createSheet("Index", 0);
            Sheet readSheet ;
            double oldTime = -1.0;
            for(int writeSheetID=0;writeSheetID<totalInputSheet;writeSheetID++)//
            {
                readSheet = wbToRead.getSheet(writeSheetID);
                for (int j = 1; j < readSheet.getRows(); j++)
                {
                    from = readSheet.getCell(0, j).getContents();
                    to = readSheet.getCell(1, j).getContents();
                    if(j == 615)
                    {
                        System.out.println("WOW");
                        j++;
                        j--;
                    }
                    freq = Double.parseDouble(readSheet.getCell(2, j).getContents());
                    if((!isInOne && j==1) || (isInOne && freq != oldTime))//first time in both cases, and each time if time changes for isInOne
                    {
                        writingRow = 1;
                        s = workbook.createSheet(timeUnit + writeSheetID, (writeSheetID+1));
                        s.addCell(new Label(0,0,"Source"));
                        s.addCell(new Label(1,0,"Target"));
                        s.addCell(new Label(2,0,"Weight"));
                        if(isInOne)
                        {
                            oldTime = freq;
                            writeSheetID++;
                            System.out.println("Processing another time");
                        }
                    }
                    
                    if(from.equals(to) && removeDuplicate)
                        continue;
                    
                    fromID = nameList.indexOf(from);
                    toID = nameList.indexOf(to);
                    if(fromID == -1)
                    {
                       nameList.add(from);
                       fromID = nameList.indexOf(from);
                    }
                    if(toID == -1)
                    {
                        if(!from.equals(to))
                        {
                            nameList.add(to);
                            toID = nameList.indexOf(to);
                        }
                        else
                        {
                            System.out.println("Avoiding repeatation");
                            toID = fromID;
                        }
                    }         
                    
                    s.addCell(new Number(0,writingRow,fromID));
                    s.addCell(new Number(1,writingRow,toID));
                    if(!isInOne)
                        s.addCell(new Number(2,writingRow,freq));
                    else
                        s.addCell(new Number(2,writingRow,1));
                    writingRow++;
                }//end traversing all rows of current Sheet
                System.out.println("finished processing sheet " + writeSheetID);

            }
            
            s = workbook.getSheet(0);
            s.addCell(new Label(0,0,"ID"));
            s.addCell(new Label(1,0,"Email"));
            for(int i=0;i<nameList.size();i++)
            {
                s.addCell(new Number(0,i+1,i));
                s.addCell(new Label(1,i+1,nameList.get(i)));
            }        
            workbook.write();
            workbook.close();
            System.out.println("Done, wow");
        } catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }
    public void gCalculateANFromSINStep2(int totalInputSheet, int startSheetOffset, String timeUnit, String filenameToRead, String fileNameToWrite)
    {
        int from, to;
        double freq;
        int firstOccurrence, lastOccurrence;
        ArrayList <Integer> fromNodeID = new ArrayList<>();//   1 1 1 3 2 2
        ArrayList <Integer> toNodeID = new ArrayList<>();//     2 3 4 4 1 2
        ArrayList <Double> freqOfOccurrence = new ArrayList<>();
        try
        {
          File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            Sheet readSheet ;  
            for(int i=0;i<totalInputSheet;i++)
            {
                readSheet = wbToRead.getSheet(i+startSheetOffset);
                for (int j = 1; j < readSheet.getRows(); j++)
                {
                    from = Integer.parseInt(readSheet.getCell(0, j).getContents());
                    to = Integer.parseInt(readSheet.getCell(1, j).getContents());
                    freq = Double.parseDouble(readSheet.getCell(2, j).getContents()); 
                    //search if from, to edge is found in two Arraylist, this means the edge is present
                    firstOccurrence = fromNodeID.indexOf(from);//should get first occurrence ID
                    if(firstOccurrence != -1)//from ID is present, search all of from ID's
                    {
                        lastOccurrence = fromNodeID.lastIndexOf(from);//now search from firstOccurrence to last for the other part of the edge
                        int k;
                        for(k=firstOccurrence;k<=lastOccurrence;k++)
                        {
                            if(toNodeID.get(k)==to)//got a match, the edge is present
                            {
                                freqOfOccurrence.set(k, freqOfOccurrence.get(k)+freq);//added the freq
                                break;
                            }
                        }
                        if(k>lastOccurrence)//we couldn't find an edge, so insert it after lastOccurrence
                        {
                            fromNodeID.add(lastOccurrence+1,from);
                            toNodeID.add(lastOccurrence+1,to);
                            freqOfOccurrence.add(lastOccurrence+1,freq);
                        }
                    }
                    else//the edge is not present. insert it at the last
                    {
                        fromNodeID.add(from);
                        toNodeID.add(to);
                        freqOfOccurrence.add(freq);
                    }
                }//end traversing rows of current sheet
                System.out.println("Finished processing " + timeUnit + " " + i);
            }//end traversing all the required sheets
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableSheet s = workbook.createSheet("Index", 0);
            
            s.addCell(new Label(0,0,"Source"));
            s.addCell(new Label(1,0,"Target"));
            s.addCell(new Label(2,0,"Weight"));
                        
            for(int i =0;i<fromNodeID.size();i++)
            {
                s.addCell(new Number(0,i+1,fromNodeID.get(i)));    
                s.addCell(new Number(1,i+1,toNodeID.get(i)));    
                s.addCell(new Number(2,i+1,freqOfOccurrence.get(i)));    
            }
            workbook.write();
            workbook.close();
            System.out.println("DONE Calculating AN");
        }//end trying read and/or write
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }        
    }
    
    /*
     * Input data must be in Sheet 0 with ID in 1st column. 1st Row is Header. Consecuting rows have data
     */
    //**Problems: The values gets rounded. Check it.
    public void mergeAllSINCentralityInOneFileForAllActorStep3(int totalActor, String inputDirectory, int startFileName, int endFileName, String fileNameToWrite, int centralityColumn, String centralityName)
    {
//        int totalInputSheet = 6;//July to December
        int currentActor;
        WritableSheet s;
        Sheet readSheet;
        File inputWorkbook;
        Workbook wbToRead;
        //int centralityCol;
        ArrayList<Double> SINCentrality = new ArrayList<>();
   //     ArrayList<Double> SINBtwnns = new ArrayList<>();
    //    ArrayList<Double> SINClsns = new ArrayList<>();
        
        try
        {   
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableCellFormat wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            s = workbook.createSheet(centralityName, 0);
            
            s.addCell(new Label(0,0,"ID"));
            for(int i=0;i<totalActor;i++)
            {
                s.addCell(new Number(0,i+1,i));
            }
            
            for(int i=startFileName;i<=endFileName;i++)//loop through ALL SIN
            {
                inputWorkbook = new File(inputDirectory+"\\"+i+".xls");
                wbToRead = Workbook.getWorkbook(inputWorkbook);
            
                //reset the Arraylist
                SINCentrality.clear();
//                SINBtwnns.clear();
  //              SINClsns.clear();
                for(int j=0;j<totalActor;j++)
                {
                    SINCentrality.add(-1.0);//-1 means emtpy
                    //SINBtwnns.add(-1.0);
                    //SINClsns.add(-1.0);
                }
                
                readSheet = wbToRead.getSheet(0); 
                for (int j=1;j<readSheet.getRows();j++)
                {
                    currentActor = Integer.parseInt(readSheet.getCell(0,j).getContents());
                    SINCentrality.set(currentActor, Double.parseDouble(readSheet.getCell(centralityColumn,j).getContents()));
                   // SINBtwnns.set(currentActor, Double.parseDouble(readSheet.getCell(3,j).getContents()));
                   // SINClsns.set(currentActor, Double.parseDouble(readSheet.getCell(4,j).getContents()));
                }
       
                s.addCell(new Label(i-startFileName+1,0,centralityName + " " + i));
                //s.addCell(new Label(BtwCol,0,"Btwnns Month " + i));
                //s.addCell(new Label(ClsCol,0,"Clsns Month " + i));
                for(int j = 0; j<totalActor;j++)//traverse all SIN, If not found they starts from sheet 2
                {
                    s.addCell(new Number(i-startFileName+1,j+1,SINCentrality.get(j),wcf));
                  //  s.addCell(new Number(BtwCol,j+1,SINBtwnns.get(j)));
                  //  s.addCell(new Number(ClsCol,j+1,SINClsns.get(j)));
                }
                System.out.println("end traversing " +inputDirectory+"\\"+i+".xls");
            }//end traversing all the sheets
            workbook.write();
            workbook.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }        
    }
    //function first required for HICCS paper experiment on student email dataset on weekly basis
    public void calculateOutDegreeForMultipleSheetsAndMergeOne(int totalSheets, int startOffset, int totalActors, String filenameToRead, String fileNameToWrite)
    {
        WritableSheet s;
        Sheet readSheet;
        File inputWorkbook;
        Workbook wbToRead;
        int fromID, freq;
        ArrayList <Integer> outDegreeList = new ArrayList<>(totalActors);
        try
        {   
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableCellFormat wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            s = workbook.createSheet("Merged Data", 0);
            
            s.addCell(new Label(0,0,"ID"));
            inputWorkbook = new File(filenameToRead);
            wbToRead = Workbook.getWorkbook(inputWorkbook);
                
            s.addCell(new Label(0,0,"ID"));
            for(int i=0;i<totalActors;i++)
            {
                s.addCell(new Number(0,i+1,i)); 
            }
            //traverse through all sheets
            for(int i = 0; i< totalSheets; i++)
            {
                //clear all actor's outgoing degree list
                outDegreeList.clear();
                for (int j=0;j<totalActors;j++)
                {
                    outDegreeList.add(j, -1);
                }
                //get the sheet
                 readSheet = wbToRead.getSheet(i+startOffset);
                 
                 //traverse through all rows of this sheet to measure outgoing degree.
                 for(int j=1;j<readSheet.getRows();j++)
                 {
                     fromID = Integer.parseInt(readSheet.getCell(0, j).getContents());
                     freq = Integer.parseInt(readSheet.getCell(2,j).getContents());
                     if(outDegreeList.get(fromID) == -1)
                         outDegreeList.set(fromID, freq);
                     else
                         outDegreeList.set(fromID, freq+outDegreeList.get(fromID));
                     
                     fromID = Integer.parseInt(readSheet.getCell(1,j).getContents());
                     if(outDegreeList.get(fromID) == -1)
                         outDegreeList.set(fromID,0);//it is present, but it was listed as absent so far. So updating.
                 }
                 //finished building the array. write all in the excel
                 s.addCell(new Label(i+1,0,readSheet.getName()));
                 for(int j=0;j<totalActors;j++)
                 {
                     s.addCell(new Number(i+1,j+1,outDegreeList.get(j)));
                 }
                 System.out.println("Done processing sheet " + i);
            }                        
            workbook.write();
            workbook.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        } 
        catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }
    //function first required for HICCS paper experiment on student email dataset on weekly basis in table 4 and 5
    //Data must be in Sheet 1. Data will start from column 1 and row 2. Row 1 is header. 1st Column is for AN. Subsequent columns are SIN
    public void calculateCentralityOverlap(int totalSINs, String filenameToRead, String fileNameToWrite)
    {
        WritableSheet s;
        Sheet readSheet;
        File inputWorkbook;
        Workbook wbToRead;     
        try
        { 
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            s = workbook.createSheet("Result Table 4 and 5", 0);
            
            s.addCell(new Label(0,0,"Top-rank list size"));
            s.addCell(new Label(1,0,"Centrality Overlap between SINs"));
            s.addCell(new Label(2,0,"Centrality Overlap between AN and SINs"));
            inputWorkbook = new File(filenameToRead);
            wbToRead = Workbook.getWorkbook(inputWorkbook);
            readSheet = wbToRead.getSheet(0);
            
            //whatever you do, start from 5 and increase by 5 upto near total actor
            for(int i = 5; i<=40;i+=5)
            {
                int totalOverlap = 0;
                //at first for table 4, between all SIN pairs.
                for(int j =0; j< totalSINs-1;j++)
                {
                    for(int k = j+1;k<totalSINs;k++)
                    {//j and k are now pairs.
                        totalOverlap += this.calculateCentralityOverlapBetweenGivenColumns(readSheet, j+1, k+1, 1, i);
                    }
                }
                //finished. Write in excel
                s.addCell(new Number(0,i/5,i));
                s.addCell(new Number(1,i/5,totalOverlap));
                System.out.println("Done table 4 for actor size " + i);
                //now for table 5, between AN and the SINs
                totalOverlap = 0;
                for(int j=0;j<totalSINs;j++)
                {
                    totalOverlap += this.calculateCentralityOverlapBetweenGivenColumns(readSheet, 0, j+1, 1, i);
                }
                s.addCell(new Number(2,i/5,totalOverlap));
                System.out.println("Done table 5 for actor size " + i);
            }
            System.out.println("Done fully");
            workbook.write();
            workbook.close();
        }
        catch (IOException e)
        {
            e.printStackTrace();
        } 
        catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
        
    }
    private int calculateCentralityOverlapBetweenGivenColumns(Sheet readSheet, int column1, int column2, int rowStart, int rowEnd)
    {
        int actorID, tempActor;
        int freqOfOccurrence = 0;
        for(int i=rowStart;i<=rowEnd;i++)
        {
            actorID = Integer.parseInt(readSheet.getCell(column1,i).getContents());
            //find that actor occurrence
            for(int j = rowStart;j<=rowEnd;j++)
            {
                tempActor = Integer.parseInt(readSheet.getCell(column2,j).getContents());
                if(actorID == tempActor)//got a match. increment Counter
                {
                    freqOfOccurrence++;
                    break;
                }
            }
        }
        return freqOfOccurrence;
    }
    /*
     * Input must be in sheet0, 1st column ID, second column AN, then all SINs. First row is header. -1 in any SIN means absent
     */
    public void calculateMesoDDAStep4(int totalSINs, String filenameToRead, String fileNameToWrite, String centralityName )
    {
        WritableSheet s;
        double OVANCentrality;
        double DDACentrality;
        boolean prevPresentCentrality;
        Sheet readSheet;
        try
        {
            File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            
            
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            s = workbook.createSheet("Index", 0);
            s.addCell(new Label(0,0,"ID"));//ID Meso1 Meso2 ... DDA
            s.addCell(new Label(totalSINs + 1, 0, "DDA " + centralityName));
            double mesoMeasure;    
            readSheet = wbToRead.getSheet(0);
            
            for(int i=1;i<readSheet.getRows();i++)//for all actors
            {
                //init
                OVANCentrality = Double.parseDouble(readSheet.getCell(1, i).getContents());
                DDACentrality = 0.0;//get ready
                
                prevPresentCentrality = true;//initially think everyone is present
                for(int j = 0; j<totalSINs;j++)//traverse all SIN
                {   
                    mesoMeasure = 0.0;
                    if(i==1)
                    {
                        s.addCell(new Label(j+1,0,"Meso " + j));
                    }
                    if(Double.parseDouble(readSheet.getCell(j+2,i).getContents()) > -1.0)//present, so consider
                    {
                        if(j==0)//first time
                        {
                            DDACentrality = DDACentrality + 1.0 * Math.abs(OVANCentrality - Double.parseDouble(readSheet.getCell(j+2,i).getContents()));
                            mesoMeasure = Math.abs(OVANCentrality - Double.parseDouble(readSheet.getCell(j+2,i).getContents()));
                        }
                        else
                        {
                            DDACentrality = DDACentrality + this.getAlpha(prevPresentCentrality, true)* Math.abs(OVANCentrality - Double.parseDouble(readSheet.getCell(j+2,i).getContents()));                            
                            mesoMeasure = this.getAlpha(prevPresentCentrality, true) * Math.abs(OVANCentrality - Double.parseDouble(readSheet.getCell(j+2,i).getContents()));
                        }
                        prevPresentCentrality = true;
                    }
                    else
                    {//absent
                        if(j == 0)//first time
                        {
                            DDACentrality = DDACentrality + 1.0 * Math.abs(OVANCentrality);
                            mesoMeasure = Math.abs(OVANCentrality);
                        }
                        else
                        {
                            DDACentrality = DDACentrality + this.getAlpha(prevPresentCentrality, false)* Math.abs(OVANCentrality);
                            mesoMeasure = this.getAlpha(prevPresentCentrality, false) * Math.abs(OVANCentrality);
                        }
                        prevPresentCentrality = false;                        
                    }
                    //write Meso
                    s.addCell(new Number(j+1,i,mesoMeasure));
                }//end traversing the SIN's
                //finalize DDA by dividing by SIN number and write
                DDACentrality = DDACentrality / totalSINs;
                
                s.addCell(new Number(0,i,i-1));
                s.addCell(new Number(totalSINs+1,i,DDACentrality));
                System.out.println("Done Actor ID " + (i-1));
            }//end traversing one actor
            workbook.write();
            workbook.close();   
            System.out.println("It is Really finsihed");
        } catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }
    

    public void formatAndExportEmailDataset(String filenameToRead, String fileNameToWrite)
    {
        int totalInputSheet = 6;//6 months
        ArrayList <String> nameList = new ArrayList<>();//for holding the names
        String from, to;
        WritableSheet s;
        int fromID, toID;
        double freq;
        
        try
        {
            File inputWorkbook = new File(filenameToRead);
            Workbook wbToRead = Workbook.getWorkbook(inputWorkbook);
            
            
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(fileNameToWrite), ws);
            WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
            WritableCellFormat cf = new WritableCellFormat(wf);
            WritableCellFormat wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            cf.setWrap(true);
            s = workbook.createSheet("Index", 0);
            Sheet readSheet ;
            
            for(int i=0;i<totalInputSheet;i++)//
            {
                readSheet = wbToRead.getSheet(i);
                s = workbook.createSheet("Month" + i, (i+1));
                s.addCell(new Label(0,0,"FROM ID",cf));
                s.addCell(new Label(1,0,"TO ID",cf));
                s.addCell(new Label(2,0,"FREQ",cf));
                
                for (int j = 1; j < readSheet.getRows(); j++)
                {
                    from = readSheet.getCell(0, j).getContents();
                    to = readSheet.getCell(1, j).getContents();
                    freq = Double.parseDouble(readSheet.getCell(2, j).getContents());
                    fromID = nameList.indexOf(from);
                    toID = nameList.indexOf(to);
                    if(fromID == -1)
                    {
                       nameList.add(from);
                       fromID = nameList.indexOf(from);
                    }
                    if(toID == -1)
                    {
                        if(!from.equals(to))
                        {
                            nameList.add(to);
                            toID = nameList.indexOf(to);
                        }
                        else
                        {
                            System.out.println("Avoiding repeatation");
                            toID = fromID;
                        }
                    }         

                    s.addCell(new Number(0,j,fromID));
                    s.addCell(new Number(1,j,toID));
                    s.addCell(new Number(2,j,freq));
                }
                System.out.println("finished processing sheet " + i);
            }
            
            s = workbook.getSheet(0);
            s.addCell(new Label(0,0,"ID",cf));
            s.addCell(new Label(1,0,"Email",cf));
            for(int i=0;i<nameList.size();i++)
            {
                s.addCell(new Number(0,i+1,i));
                s.addCell(new Label(1,i+1,nameList.get(i),cf));
            }        
            workbook.write();
            workbook.close();
            System.out.println("Done, wow");
        } catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }  
        catch (BiffException e)
        {
            e.printStackTrace();
        }
    }
    
    public void readAndPopulateMyGraph(MyDirectedGraph dg,int sheetNumber) throws IOException
    {
        File inputWorkbook = new File(inputFile);
        int minAttribute = Integer.MAX_VALUE;
        int maxAttribute = Integer.MIN_VALUE;
        Workbook w;
        try
        {
            w = Workbook.getWorkbook(inputWorkbook);
            // Get the first sheet
            Sheet sheet = w.getSheet(sheetNumber);

            ArrayList<MyNode> al = new ArrayList();
            for (int i = 0; i < sheet.getRows(); i++)
            {

                Cell x, y, z;
                x = sheet.getCell(0, i);
                y = sheet.getCell(1, i);
                z = sheet.getCell(2, i);
                if (x.getType() != CellType.NUMBER || y.getType() != CellType.NUMBER || z.getType() != CellType.NUMBER)
                {
                    System.out.println("Invalid Data at Row " + i);
                } else
                {
                    int from = Integer.parseInt(x.getContents());
                    int to = Integer.parseInt(y.getContents());
                    int attr = Integer.parseInt(z.getContents());
                    if(attr < minAttribute)
                        minAttribute = attr;
                    else if(attr > maxAttribute)
                        maxAttribute = attr;
                    dg.addNodeInNodeList(from, 1.0f, true);
                    dg.addNodeInNodeList(to, 1.0f, false);
                    dg.addAttributedEdge(attr, from, to, 1.0f);
                }
            }
            
            //dg.printGraph();
            dg.setMaxAttribute(maxAttribute);
            dg.setMinAttribute(minAttribute);

            /*
             for (int j = 0; j < sheet.getColumns(); j++) {
             for (int i = 0; i < sheet.getRows(); i++) {
             Cell cell = sheet.getCell(j, i);
             CellType type = cell.getType();
             if (type == CellType.LABEL) {
             System.out.println("I got a label "
             + cell.getContents());
             }

             if (type == CellType.NUMBER) {
             System.out.println("I got a number "
             + cell.getContents());
             }
             }
             }*/
        } catch (BiffException e)
        {
            e.printStackTrace();
        }
    }            
    public void exportJTableModelDataToExcel(String filename, DefaultTableModel dataModel)
    {
        try
        {
            // Writing Data to ExcelSheet
            WorkbookSettings ws = new WorkbookSettings();
            ws.setLocale(new Locale("en", "EN"));
            WritableWorkbook workbook = Workbook.createWorkbook(new File(filename), ws);
            WritableSheet s = workbook.createSheet("DataSheet", 0);
            this.writeData(s, dataModel);
            
            workbook.write();
            workbook.close();
        } catch (IOException e)
        {
            e.printStackTrace();
        } catch (WriteException e)
        {
            e.printStackTrace();
        }
    }

    private void writeData(WritableSheet s, DefaultTableModel dataModel) throws WriteException
    {
        WritableFont wf = new WritableFont(WritableFont.ARIAL, 10, WritableFont.BOLD);
        WritableCellFormat cf = new WritableCellFormat(wf);
        WritableCellFormat wcf;
        cf.setWrap(true);
        
        for(int i =0;i<dataModel.getColumnCount();i++)
        {
            //if(dataModel.getColumnName(i).compareToIgnoreCase(inputFile))
            wcf = new WritableCellFormat(NumberFormats.DEFAULT);
            s.addCell(new Label(i,0,dataModel.getColumnName(i),cf));
            for(int j=0;j<dataModel.getRowCount();j++)
            {
                s.addCell(new Number(i,j+1,Double.parseDouble(dataModel.getValueAt(j, i).toString())));    
            }
        }
        if(true)
            return;
        /* Creates Label and writes date to one cell of sheet */
        Label l = new Label(0, 0, "Date", cf);
        s.addCell(l);
        WritableCellFormat cf1 = new WritableCellFormat(DateFormats.FORMAT9);

        DateTime dt = new DateTime(0, 1, new Date(), DateTime.GMT);

        s.addCell(dt);

        /* Creates Label and writes float number to one cell of sheet */
        l = new Label(2, 0, "Float", cf);
        s.addCell(l);
        WritableCellFormat cf2 = new WritableCellFormat(NumberFormats.FLOAT);
        jxl.write.Number n = new jxl.write.Number(2, 1, 3.1415926535, cf2);
        s.addCell(n);

        n = new jxl.write.Number(2, 2, -3.1415926535, cf2);
        s.addCell(n);

        /* 
         * Creates Label and writes float number upto 3 decimal to one cell of 
         * sheet 
         */
        l = new Label(3, 0, "3dps", cf);
        s.addCell(l);
        NumberFormat dp3 = new NumberFormat("#.###");
        WritableCellFormat dp3cell = new WritableCellFormat(dp3);
        n = new jxl.write.Number(3, 1, 3.1415926535, dp3cell);
        s.addCell(n);

        /* Creates Label and adds 2 cells of sheet */
        l = new Label(4, 0, "Add 2 cells", cf);
        s.addCell(l);
        n = new jxl.write.Number(4, 1, 10);
        s.addCell(n);
        n = new jxl.write.Number(4, 2, 16);
        s.addCell(n);
        Formula f = new Formula(4, 3, "E2+E3");
        s.addCell(f);

        /* Creates Label and multipies value of one cell of sheet by 2 */
        l = new Label(5, 0, "Multipy by 2", cf);
        s.addCell(l);
        n = new jxl.write.Number(5, 1, 10);
        s.addCell(n);
        f = new Formula(5, 2, "F2 * 3");
        s.addCell(f);

        /* Creates Label and divide value of one cell of sheet by 2.5 */
        l = new Label(6, 0, "Divide", cf);
        s.addCell(l);
        n = new jxl.write.Number(6, 1, 12);
        s.addCell(n);
        f = new Formula(6, 2, "F2/2.5");
        s.addCell(f);

        /*-----------------------------------------------------------*/
        /* Format the Font */
        WritableFont wf3 = new WritableFont(WritableFont.TIMES, 10,
                WritableFont.BOLD);
        WritableCellFormat cf3 = new WritableCellFormat(wf3);
        cf3.setWrap(true);

        l = new Label(2, 8, "Nagesh", cf3);
        s.addCell(l);

        l = new Label(1, 9, "Water", cf3);
        s.addCell(l);
        n = new jxl.write.Number(2, 9, 35);
        s.addCell(n);

        l = new Label(1, 10, "Elictricity", cf3);
        s.addCell(l);
        n = new jxl.write.Number(2, 10, 57);
        s.addCell(n);

        l = new Label(1, 11, "Rent", cf3);
        s.addCell(l);
        n = new jxl.write.Number(2, 11, 750);
        s.addCell(n);

        f = new Formula(2, 13, "C10+C11+C12");
        s.addCell(f);
    }
}
