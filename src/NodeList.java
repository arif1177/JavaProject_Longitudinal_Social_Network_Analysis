
import java.util.ArrayList;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Arif Khan
 */
public class NodeList
{
    private ArrayList<MyNode> nodeList;
    public NodeList()
    {
        nodeList = new ArrayList();
    }
    public void clearNodeList()
    {
        this.nodeList.clear();
    }
    public int getNodeCount()
    {
        return nodeList.size();
    }
    /**
     * Returns the ID of the node
     * @param id The index number of the node
     * @return the node ID
     */
    public int getNodeIDAt(int id)
    {
        return this.nodeList.get(id).getNodeId();
    }
    public String getNodeLabelAt(int id)
    {
        return Integer.toString(this.nodeList.get(id).getNodeId());
    }
    public MyNode getNodeAt(int id)
    {
        return this.nodeList.get(id);
    }
    public void addNode(int nodeID, float weight, boolean isFromNode)
    {
        //if the Node is already there, increase outDegree, else add with new weight
        for(int i = 0;i<nodeList.size();i++)
        {
            if(nodeList.get(i).getNodeId()== nodeID)//this is already in the nodeList, increase the outDegree
            {   
                if(isFromNode)
                {
                    nodeList.get(i).setOutDegree(weight+nodeList.get(i).getOutDegree());
                }
                else
                {
                    nodeList.get(i).setInDegree(weight+nodeList.get(i).getInDegree());
                }
                return;
            }
        }
        //Should create and add a New Node
        if(isFromNode)
        {
            nodeList.add(new MyNode(0.0f,weight,nodeID));
        }
        else
        {
            nodeList.add(new MyNode(weight,0.0f,nodeID));
        }
    }
    @Override
    public String toString()
    {
        String s = "\n=========\nPrinting Node List\n";
        for(int i=0;i<this.nodeList.size();i++)
        {
            s = s+ "Node: " + this.nodeList.get(i).getNodeId() + ", OutDegree: " + this.nodeList.get(i).getOutDegree()+ ", InDegree: " + this.nodeList.get(i).getInDegree()+ "\n";
        }   
        return s;
    }
}
