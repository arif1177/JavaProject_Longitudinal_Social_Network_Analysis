import javax.swing.table.DefaultTableModel;
import org.gephi.graph.api.DirectedGraph;
import org.gephi.ranking.api.Ranking;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author Arif Khan
 */
public class DataTableController
{

   private      MyDataModel dataModel;
//private      DefaultTableModel dataModel;
    public DataTableController()
    {
//        dataModel = new MyDataModel(new Object[]
//                {
//                    "ID", "Label", "OutDegree"
//                }, 0);
        Object [][] data = {{"0","0",0.0f}};
        String []columns = {"ID", "Label", "OutDegree"};
           dataModel = new MyDataModel(data,columns);
    }
    /**
     * Returns the table model
     *
     * @return <code>DefaultTableModel</code>, that is used to hold Node data
     */
    public DefaultTableModel getDataModel()
    {
        return this.dataModel;
    }

    /**
     * Clears the data table. Only keeps 3 basic columns, and then inserts basic
     * node data from
     * <code>mdg</code> whose attribute are within
     * <code>minAttribute</code> and
     * <code>maxAttribute</code>
     *
     * @param mdg Has the necessary information on all the graph node, edges,
     * property.
     * @param minAttribute Minimum attribute value to consider for Node data
     * calculation
     * @param maxAttribute Maximum attribute value to consider for Node data
     * calculation
     */
    public void clearThenAddBasicGraphNode(MyDirectedGraph mdg, int minAttribute, int maxAttribute)
    {
        NodeList n = mdg.getNodeList();
        dataModel.setColumnCount(3);//the remaining columns are declared unnecessary and will not be considered
        for (int i = dataModel.getRowCount() - 1; i >= 0; i--)//loop to delete all the rows
        {
            dataModel.removeRow(i);
        }

        for (int i = 0; i < n.getNodeCount(); i++)//loop through all the nodes. Find data(e.g. Out Degree)
        {
            float f = 0.0f;
            for (int j = 0; j < n.getNodeCount(); j++)
            {
                if (true)//not skipping self edge
                {
                    f = f + mdg.findTotalEdgeWeight(n.getNodeIDAt(i), n.getNodeIDAt(j), minAttribute, maxAttribute);
                }
            }
            dataModel.addRow(new Object[]
                    {
                        //n.getNodeIDAt(i) + "", n.getNodeIDAt(i) + "", f + ""
                n.getNodeIDAt(i) + "", n.getNodeIDAt(i) + "", f
                    });//new row added.
        }
    }

    /**
     * Add a new ranking column data in Data Table depending on the parameters.
     * If the column is already present update all the rows.
     *
     * @param name Ranking name i.e. Degree Centrality, Betweenness centrality
     * etc. We search the existing columns with this name to check if it is
     * already present or not
     * @param ranking Class that has the ranking data
     * @param directedGraph Gephi graph instance
     * @param n List of the Nodes.
     */
    public void addNewRankingColumn(String name, Ranking ranking, DirectedGraph directedGraph, NodeList n)
    {

        int j = dataModel.getColumnCount();
        for (int i = 0; i < j; i++)//check if the column name is already there. If yes then just update all the rows.
        {
            if (dataModel.getColumnName(i).compareTo(name) == 0)//check if matches
            {
                j = i;//Found the column. Just point to it.
                break;
            }
        }
        if (j == dataModel.getColumnCount())//insert only if column not found
        {
            dataModel.addColumn(name + "");
        }
        int i;
        for (i = 0; i < directedGraph.getNodeCount(); i++)
        {
            //dataModel.setValueAt( ranking.getValue(directedGraph.getNode(i))+"", i-1,j); 
            try
            {
            dataModel.setValueAt(ranking.getValue(directedGraph.getNode(n.getNodeLabelAt(i))), i, j);
            }
               catch(Exception ex)
        {
            ex.printStackTrace();
        }
        }
    }
}

