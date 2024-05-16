using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ExcelDna.Integration.CustomUI;

namespace Test18
{
    [ComVisible(true)]
    public class Ribbon : ExcelRibbon
    {
        IRibbonUI _ribbon;

        public override string GetCustomUI(string RibbonID)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnLoad'>
  <ribbon>
    <tabs>
      <tab id='MyTab' label='TestComboBox'>
        <group id='MyGroup' label='My Group'>
          <comboBox id='MyComboBox' 
                    getItemCount='GetComboBoxItemCount'
                    getItemLabel='GetComboBoxItemLabel'
                    getText='GetComboBoxText'
                    onChange='OnComboBoxTextChanged' />
          <button id='MyButton' label='Add Item' size='large' onAction='OnAddItemClicked' />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }

        // Implement OnLoad callback
        public void OnLoad(IRibbonUI ribbon)
        {
            _ribbon = ribbon;
            _comboBoxWrapper = new RibbonComboBoxWrapper(_ribbon, "MyComboBox");
            _comboBoxWrapper.TextChanged += ComboBoxWrapper_TextChanged;
            
            // Example usage
            _comboBoxWrapper.AddItem("Item 1");
            _comboBoxWrapper.AddItem("Item 2");
            _comboBoxWrapper.SetText("Select an item");
        }

        private void ComboBoxWrapper_TextChanged(object sender, string newText)
        {
            // Handle text changed event
            MessageBox.Show($"ComboBox text changed to: {newText}");
        }

        // Implement OnAddItemClicked callback
        public void OnAddItemClicked(IRibbonControl control)
        {
            _comboBoxWrapper.AddItem($"Item {MyComboBox.GetItemCount(control) + 1}");
        }

        #region Combobox wrapper
        public RibbonComboBoxWrapper _comboBoxWrapper;
        public RibbonComboBoxWrapper MyComboBox => _comboBoxWrapper;


        public int GetComboBoxItemCount(IRibbonControl control)
        {
            return _comboBoxWrapper.GetItemCount(control);
        }

        public string GetComboBoxItemLabel(IRibbonControl control, int index)
        {
            return _comboBoxWrapper.GetItemLabel(control, index);
        }

        public string GetComboBoxText(IRibbonControl control)
        {
            return _comboBoxWrapper.GetComboBoxText(control);
        }

        public void OnComboBoxTextChanged(IRibbonControl control, string text)
        {
            _comboBoxWrapper.OnComboBoxTextChanged(control, text);
        }
        #endregion

    }

    public class RibbonComboBoxWrapper
    {
        private IRibbonUI _ribbon;
        private string _controlId;

        private List<string> items = new List<string>();
        private string text;
        public event EventHandler<string> TextChanged;

        public RibbonComboBoxWrapper(IRibbonUI ribbon, string controlId)
        {
            _ribbon = ribbon;
            _controlId = controlId;
        }

        public void AddItem(string item)
        {
            items.Add(item);
            InvalidateControl();
        }

        public void SetItems(List<string> newItems)
        {
            items = newItems;
            InvalidateControl();
        }

        public void SetText(string newText)
        {
            text = newText;
            InvalidateControl();
        }

        public string GetText()
        {
            return text;
        }

        private void InvalidateControl()
        {
            // Invalidate the ComboBox control to force the Ribbon to call the callbacks again
            _ribbon.InvalidateControl(_controlId);
        }

        public int GetItemCount(IRibbonControl control)
        {
            return items.Count;
        }

        public string GetItemLabel(IRibbonControl control, int index)
        {
            return items[index];
        }

        public string GetComboBoxText(IRibbonControl control)
        {
            return text;
        }

        public void OnComboBoxTextChanged(IRibbonControl control, string newText)
        {
            text = newText;
            TextChanged?.Invoke(this, newText);
        }
    }
}
