using FlaUI.Core.AutomationElements;

namespace PowerGUIAutomation;

// The red X (circle, #F44336) and the green V (#4CAF50) that the GUI shows to the right of a parameter's value field are icons that UI Automation does
// NOT expose, so they are found by their COLOURS in a screenshot of the area right of the field (the field's row only). A few dozen pixels of the colour
// = the icon is there. This needs Windows display scaling of 100% and the GUI window fully visible (not minimized, covered or off the screen).
// Used by every test that writes values (the same rule everywhere): a red X = the unit did not accept the value.
public static class RowIcons
{
    public static (bool RedX, bool GreenV) Read(AutomationElement valueEl)
    {
        // the ListBoxItem of the row is two levels above the value field
        var item = valueEl.Parent?.Parent;
        return item == null ? (false, false) : Read(valueEl, item);
    }

    public static (bool RedX, bool GreenV) Read(AutomationElement valueEl, AutomationElement item)
    {
        try
        {
            var field = valueEl.BoundingRectangle;
            var row = item.BoundingRectangle;
            int x = (int)field.Right;
            int width = (int)Math.Min(row.Right - field.Right, 320);
            if (width < 10 || row.Height < 10)
            {
                return (false, false);
            }

            using var shot = FlaUI.Core.Capturing.Capture.Rectangle(new System.Drawing.Rectangle(x, (int)row.Top, width, (int)row.Height));
            var bmp = shot.Bitmap;
            int red = 0, green = 0;
            for (int py = 0; py < bmp.Height; py++)
            {
                for (int px = 0; px < bmp.Width; px++)
                {
                    var c = bmp.GetPixel(px, py);
                    if (c.R >= 200 && c.G <= 100 && c.B <= 100 && c.R - c.G >= 110)
                    {
                        red++;
                    }
                    else if (c.G >= 140 && c.R <= 120 && c.B <= 140 && c.G - c.R >= 50)
                    {
                        green++;
                    }
                }
            }

            return (red >= 25, green >= 25);
        }
        catch
        {
            return (false, false);
        }
    }
}
