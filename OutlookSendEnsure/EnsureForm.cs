using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace OutlookSendEnsure
{
    public partial class EnsureForm : Form
    {
        public int buttonIndex = 0;
        public List<object> addressList = new List<object>(); // 存储组邮件地址对象
                                                              // 普通组地址对象类型为 Outlook AddressEntries
                                                              // 自建组地址对象类型为 Outlook DistListItem
        private Dictionary<string, Dictionary<string,string>> i18n = new Dictionary<string, Dictionary<string, string>>(); // 国际化字典
        private string lang = "ch";

        public const int TO = 1;
        public const int CC = 2;


        public EnsureForm() // 构造函数
        {
            InitializeComponent(); // 默认构造函数

            // 构造多语言环境
            Dictionary<string, string> ch = new Dictionary<string, string>();
            ch.Add("formTitle", "再次确认收件和抄送地址。");
            ch.Add("mailAddress", "邮件地址");
            ch.Add("groupMailAddress", "组邮件地址");
            ch.Add("ok", "确认");
            ch.Add("cancel", "取消");
            ch.Add("to", "收件");
            ch.Add("cc", "抄送");
            this.i18n.Add("ch", ch);

            Dictionary<string, string> en = new Dictionary<string, string>();
            en.Add("formTitle", "Please make a double check on receiver and cc address.");
            en.Add("mailAddress", "Mail Address");
            en.Add("groupMailAddress", "Group Mail Address");
            en.Add("ok", "OK");
            en.Add("cancel", "Cancel");
            en.Add("to", "To");
            en.Add("cc", "CC");
            this.i18n.Add("en", en);

            Dictionary<string, string> jp = new Dictionary<string, string>();
            jp.Add("formTitle", "宛先とCCをもう一度確認してください。");
            jp.Add("mailAddress", "メールア ドレス");
            jp.Add("groupMailAddress", "グループ メールア  ドレス");
            jp.Add("ok", "確認する");
            jp.Add("cancel", "キャンセル");
            jp.Add("to", "宛先");
            jp.Add("cc", "Cc");
            this.i18n.Add("jp", jp);

            this.Text = this.i18n[this.lang]["formTitle"];
            this.label1.Text = this.i18n[this.lang]["to"];
            this.label2.Text = this.i18n[this.lang]["cc"];
            this.button1.Text = this.i18n[this.lang]["ok"];
            this.button2.Text = this.i18n[this.lang]["cancel"];
        }

        public void showAddress(string searchName, Outlook.AddressEntry address, int flag){
            if (address != null) {

                // 邮件地址
                if (address.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeUserAddressEntry)  // 普通地址
                {
                    // 普通地址的 searchName 中缺少Email地址 需要另外获取
                    this.showSingle(this.i18n[this.lang]["mailAddress"] + ": " + searchName + " " + address.GetExchangeUser().PrimarySmtpAddress, flag);
                }
                else if (address.AddressEntryUserType == Outlook.OlAddressEntryUserType.olOutlookContactAddressEntry) // 存入联系人中的地址
                {
                    // 存入联系人中的地址的 searchName 中包含了Email地址 无要另外获取
                    this.showSingle(this.i18n[this.lang]["mailAddress"] + ": " + searchName, flag);
                }
                // 组邮件地址
                else if (address.AddressEntryUserType == Outlook.OlAddressEntryUserType.olExchangeDistributionListAddressEntry) // 普通组地址
                {
                    // 普通组地址的 searchName 为组名 Email地址可通过Outlook ExchangeDistributionList 对象获取
                    Outlook.ExchangeDistributionList disList = address.GetExchangeDistributionList();
                    this.showMuilt(this.i18n[this.lang]["groupMailAddress"] + ": " + searchName + " " + disList.PrimarySmtpAddress, flag);
                    this.addAddressNativeGroupList(disList);
                }
                else if (address.AddressEntryUserType == Outlook.OlAddressEntryUserType.olOutlookDistributionListAddressEntry) // 自建组地址
                {
                    // 自建组地址的 searchName 为组名 且无Email地址 直接仅显示组名
                    this.showMuilt(this.i18n[this.lang]["groupMailAddress"] + ": " + searchName, flag);
                    this.addAddressSelfBuildGroupList(address, searchName);
                }
            }
            else {
                this.showSingle(searchName, 0);
            } 
        }

        public void showSingle(string text, int flag) // 显示邮件地址
        {
            TreeNode node = new TreeNode();
            node.Text = text;
            node.Name = "-1";
            if (flag == TO)
            {
                this.treeView1.Nodes.Add(node);

            }
            else if (flag == CC)
            {
                this.treeView2.Nodes.Add(node);
            }
        }

        public void showMuilt(string text, int flag) // 显示组邮件地址
        {
            TreeNode node = new TreeNode();
            node.Text = text;
            node.Name = "" + this.buttonIndex;
            node.ForeColor = Color.CornflowerBlue;
            if (flag == TO)
            {
                this.treeView1.Nodes.Add(node);

            }
            else if (flag == CC)
            {
                this.treeView2.Nodes.Add(node);
            }
            this.buttonIndex += 1;
        }

        public void addAddressNativeGroupList(Outlook.ExchangeDistributionList disList) // 普通组邮件成员list添加到addressList区存储
        {
            Outlook.AddressEntries addressEntries = disList.GetExchangeDistributionListMembers();
            this.addressList.Add(addressEntries);
        }

        public void addAddressSelfBuildGroupList(Outlook.AddressEntry address, String groupName) // 自建组列表成员list添加到addressList区存储
        {
            Outlook.NameSpace nameSpace = address.Application.GetNamespace("MAPI");
            Outlook.MAPIFolder mAPIFolder = nameSpace.GetDefaultFolder(Outlook.OlDefaultFolders.olFolderContacts);
            for (int i = 0; i < mAPIFolder.Items.Count; i++)
            {
                dynamic dyna = mAPIFolder.Items.GetNext();
                Outlook.DistListItem distListItem = (Outlook.DistListItem)dyna;
                if (groupName == distListItem.DLName)
                {
                    this.addressList.Add(distListItem);
                }
                break;
            }
        }

        private void button1_Click(object sender, EventArgs e) // 确认发送
        {
            ThisAddIn._cancel = false;
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e) // 取消发送
        {
            ThisAddIn._cancel = true;
            this.Close();
        }

        public string[] getAddresses(int index) {
            string[] addressesArr = null; // 保存邮件组下的成员地址
            if (this.addressList[index] is Outlook.DistListItem)
            {  
                Outlook.DistListItem distListItem = (Outlook.DistListItem)this.addressList[index];
                addressesArr = new string[distListItem.MemberCount];
                for (int i = 0; i < distListItem.MemberCount; i++)
                {
                    Outlook.Recipient recipient = distListItem.GetMember(i + 1);
                    addressesArr[i] = recipient.Name + " " + recipient.Address;
                }
            }
            else if (this.addressList[index] is Outlook.AddressEntries)
            {
                Outlook.AddressEntries addresses = (Outlook.AddressEntries)this.addressList[index];
                addressesArr = new string[addresses.Count];
                for (int i = 1; i <= addresses.Count; i++)
                {
                    Outlook.AddressEntry membAddress = addresses[i];
                    addressesArr[i - 1] = membAddress.Name + " " + membAddress.GetExchangeUser().PrimarySmtpAddress;

                }
            }
            return addressesArr;
        }

        private void treeView1_AfterSelect(object sender, TreeViewEventArgs e)
        {
            int nodeCount = this.treeView1.SelectedNode.Nodes.Count; // 当前选中的邮件组下的成员总数
            TreeNode selectNode = this.treeView1.SelectedNode;
            int nodeIndex = -1;
            if (nodeCount > 0) // 如果已经展开过 则返回
            {
                return;
            }
            try { nodeIndex =  Convert.ToInt32(this.treeView1.SelectedNode.Name); } catch { return; } // 如果不是邮件组 则返回
                                                                                                           // 当前选中的邮件组Name存放的是string类型的Index信息
            if (nodeIndex == -1)
            {
                return;
            }
            
            Thread th = new Thread(() => // 多线程获得数据
            {
                string [] addressesArr = this.getAddresses(nodeIndex); // 当前选中的邮件组下的成员地址

                this.BeginInvoke(new Action(() => // 主线程更新UI
                {
                    for (int i=0; i< addressesArr.Length; i+=1)
                    {
                        TreeNode node = new TreeNode();
                        node.Text = addressesArr[i];
                        selectNode.Nodes.Add(node);
                    }

                }));
            });
            th.Start(); // 启动线程 
        }

        private void treeView2_AfterSelect(object sender, TreeViewEventArgs e)
        {
            int nodeCount = this.treeView2.SelectedNode.Nodes.Count;
            TreeNode selectNode = this.treeView2.SelectedNode;
            int nodeIndex = -1;
            if (nodeCount > 0)
            {
                return;
            }
            try { nodeIndex = Convert.ToInt32(this.treeView2.SelectedNode.Name); } catch { return; }

            if (nodeIndex == -1)
            {
                return;
            }

            Thread th = new Thread(() => 
            {
                string[] addressesArr = this.getAddresses(nodeIndex);

                this.BeginInvoke(new Action(() => 
                {
                    for (int i = 0; i < addressesArr.Length; i += 1)
                    {
                        TreeNode node = new TreeNode();
                        node.Text = addressesArr[i];
                        selectNode.Nodes.Add(node);
                    }
                }));
            });
            th.Start();
        }
    }
}
