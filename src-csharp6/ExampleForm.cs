using System.Windows.Forms;

namespace Example
{
    public partial class ExampleForm : Form
    {
        public ExampleForm()
        {
            InitializeComponent();

            foreach (var printerInfo in PrinterInfo.GetInstalledPrinterNamesAndImages(imageList1.ImageSize))
            {
                imageList1.Images.Add(printerInfo.Image);

                listView1.Items.Add(new ListViewItem
                {
                    Text = printerInfo.DisplayName,
                    ImageIndex = imageList1.Images.Count - 1,
                    Tag = printerInfo
                });
            }
        }
    }
}
