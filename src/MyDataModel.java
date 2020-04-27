
import javax.swing.table.DefaultTableModel;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author User
 */
public class MyDataModel extends DefaultTableModel
{

    MyDataModel(Object[][] object, String[] i)
    {
        super(object,i);
    }
    @Override
    public Class getColumnClass(int columnIndex) {
 //       return getValueAt(0, columnIndex).getClass();
//        if(columnIndex < 2)//string for Label and ID
//            return String.class;
//       else//float for outDegree, Degree, closeness, betweenness etc.
//            return float.class;
        if(columnIndex <2)
            return String.class;
        else
            return Float.class;
    }
}
