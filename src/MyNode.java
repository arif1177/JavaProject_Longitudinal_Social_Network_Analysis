/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author User
 */
//**delete if not needed
public class MyNode
{
    private float outDegree;
    private float inDegree;
    private int nodeId;
   
    
    
    public MyNode(float inDegree, float outDegree, int nodeId)
    {
        this.inDegree = inDegree;
        this.outDegree = outDegree;
        this.nodeId = nodeId;
    }

    public float getInDegree()
    {
        return inDegree;
    }

    public float getOutDegree()
    {
        return outDegree;
    }

    public int getNodeId()
    {
        return nodeId;
    }

    public void setOutDegree(float outDegree)
    {
        this.outDegree = outDegree;
    }

    public void setInDegree(float inDegree)
    {
        this.inDegree = inDegree;
    }
    
}