
import javax.swing.InputVerifier;
import javax.swing.JComponent;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JTextField;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *
 * @author User
 */
public class MyInputVerifier extends InputVerifier
{

    @Override
    public boolean verify(JComponent input)
    {
        try
        {
            String str = ((JTextField) input).getText();
            if(str == null || str.isEmpty())
                return true;
            Integer.parseInt(str);
            return true;
        } catch (NumberFormatException e)
        {
            return false;
        }
    }
    public boolean checkRange(int minAllowedValue, int maxAllowedValue, int minValue, int maxValue)
    {
        if(minValue < minAllowedValue || maxValue > maxAllowedValue || maxValue < minValue)
        {
            JOptionPane.showMessageDialog(null, "Value is outside of allowable range. Try again", "Invalid input", JOptionPane.INFORMATION_MESSAGE );
            return false;
        }
        else
        {
            return true;   
        }
    }
}
