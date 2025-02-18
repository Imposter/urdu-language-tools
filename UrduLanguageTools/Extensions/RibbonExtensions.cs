using System;
using System.Collections.Generic;
using Microsoft.Office.Tools.Ribbon;

namespace UrduLanguageTools.Extensions
{
    public static class RibbonExtensions
    {
        public static void SetItems<T>(this RibbonDropDown dropDown, RibbonFactory factory, IEnumerable<T> items, Func<T, string> labelSelector)
        {
            dropDown.Items.Clear();
            foreach (var item in items)
            {
                var dropDownItem = factory.CreateRibbonDropDownItem();
                dropDownItem.Label = labelSelector(item);
                dropDownItem.Tag = item;

                dropDown.Items.Add(dropDownItem);
            }
        }

        public static void SetItems<T>(this RibbonComboBox comboBox, RibbonFactory factory, IEnumerable<T> items, Func<T, string> labelSelector)
        {
            comboBox.Items.Clear();
            foreach (var item in items)
            {
                var comboBoxItem = factory.CreateRibbonDropDownItem();
                comboBoxItem.Label = labelSelector(item);
                comboBoxItem.Tag = item;

                comboBox.Items.Add(comboBoxItem);
            }
        }
    }
}
