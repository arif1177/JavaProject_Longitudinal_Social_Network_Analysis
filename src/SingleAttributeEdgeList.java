
import java.util.ArrayList;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author Arif Khan
 */
public class SingleAttributeEdgeList
{
    private int attributeID;
    private ArrayList<Float> edgeWeights;
    private ArrayList<Integer> fromNodeIDs;
    private ArrayList<Integer> toNodeIDs;
    
    public SingleAttributeEdgeList(int attributeID)
    {
        this.attributeID = attributeID;
        this.edgeWeights = new ArrayList<>();
        this.fromNodeIDs = new ArrayList<>();
        this.toNodeIDs = new ArrayList<>();
    }

    public int getAttributeID()
    {
        return attributeID;
    }
    
    public void addAttributedEdge(int fromNodeID, int toNodeID, float weight)
    {
        //search in all the connected toNodes
        for(int i = 0;i<this.fromNodeIDs.size();i++)
        {
            if(this.fromNodeIDs.get(i) == fromNodeID && this.toNodeIDs.get(i) == toNodeID)
            {
                this.edgeWeights.set(i,this.edgeWeights.get(i) + weight);//same edge again, so increasing weight/frequency
                return;
            }
        }
        //this is the new edge. Let's add
        this.fromNodeIDs.add(fromNodeID);
        this.toNodeIDs.add(toNodeID);
        this.edgeWeights.add(weight);//adding the weight for newly created toNode
    }
    /**
     * 
     * @param outgoingNodeID  The outgoing nodeID
     * @param incomingNodeID The incident nodeID
     * @return Total edge weight for this attribute from outgoingNodeID to incomingNodeID
     */
    public float getToEdgeWeight(int outgoingNodeID, int incomingNodeID)
    {
       float f = 0.0f;
       for(int i=0;i<this.fromNodeIDs.size();i++)
       {
            if(this.fromNodeIDs.get(i) == outgoingNodeID && this.toNodeIDs.get(i) == incomingNodeID)
            {
                f += this.edgeWeights.get(i);//only one match possible. so return
                return f;
            }
       }
       return f;
    }
    
    
    @Override
    public String toString()
    {
        String s = "\n=========\nPrinting Edge List of Attribute " + this.attributeID + "\n";
        for(int i=0;i<this.fromNodeIDs.size();i++)
        {
            s = s + this.fromNodeIDs.get(i).toString() + " --> " + this.toNodeIDs.get(i).toString() + ". Weight: " + this.edgeWeights.get(i) + "\n";
        }                
        return s;
    }
}
