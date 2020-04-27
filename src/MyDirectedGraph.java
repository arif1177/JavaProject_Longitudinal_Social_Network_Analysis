
import java.util.ArrayList;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Arif Khan
 */
public class MyDirectedGraph
{

    private  NodeList nodeList;
    private  ArrayList<SingleAttributeEdgeList> saEdgeList;
    private int minAttribute;
    private int maxAttribute;

    public MyDirectedGraph()
    {
        nodeList = new NodeList();
        saEdgeList = new ArrayList<>();
    }
    public void addNodeInNodeList(int nodeID, float weight, boolean isFromNode)
    {
        nodeList.addNode(nodeID, weight, isFromNode);
    }
    /**
     * Adds an edge with particular attribute. If it is already there add the weight, otherwise creates the edge.
     * @param attributeID The attribute of the edge.
     * @param fromNodeID The node from which the edge is generated
     * @param toNodeID The node on which the edge is incident
     * @param weight The weight of the edge
     */
    public void addAttributedEdge(int attributeID, int fromNodeID, int toNodeID, float weight)
    {
        //if the attributeID is not present, make a new attributeList
        for(int i = 0;i<saEdgeList.size();i++)
        {
            if(saEdgeList.get(i).getAttributeID() == attributeID)//the attribute is already present, just add the edge
            {
                saEdgeList.get(i).addAttributedEdge(fromNodeID, toNodeID, weight);
                return;
            }
        }
        //the attribute is not in the list, yet. Make a new one and insert
        SingleAttributeEdgeList temp = new SingleAttributeEdgeList(attributeID);
        temp.addAttributedEdge(fromNodeID, toNodeID, weight);
        saEdgeList.add(temp);
    }

    public NodeList getNodeList()
    {
        return nodeList;
    }

    public float findTotalEdgeWeight(int fromNodeID, int toNodeID, int activeMinAttribute, int activeMaxAttribute)
    {
        float f = 0.0f;
        for (SingleAttributeEdgeList sael : saEdgeList)
        {
            if(sael.getAttributeID()>= activeMinAttribute && sael.getAttributeID()<=activeMaxAttribute)
                f += sael.getToEdgeWeight(fromNodeID, toNodeID);
        }
        return f;
    }

    public ArrayList<SingleAttributeEdgeList> getSaEdgeList()
    {
        return saEdgeList;
    }
    
    public void printGraph()
    {
        System.out.println(nodeList);
        for(SingleAttributeEdgeList temp:saEdgeList)
        {
            System.out.println(temp);
        }
    }

    public void setMinAttribute(int minAttribute)
    {
        this.minAttribute = minAttribute;
    }

    public  void setMaxAttribute(int maxAttribute)
    {
        this.maxAttribute = maxAttribute;
    }

    public  int getMinAttribute()
    {
        return this.minAttribute;
    }

    public  int getMaxAttribute()
    {
        return this.maxAttribute;
    }
}
