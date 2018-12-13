using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aras.IOM;
using System.Windows.Forms;

namespace sgm_aras_plugin
{
    public class sgmUtils
    {
        public Item moreApproval(Item Act,Item Assignment)  //act为当前节点,加签节点显示在当前节点上方100的位置，MAP需要预留出一定的空间给加签节点
        {
            Innovator inn = Act.getInnovator();
            Item WFP = inn.newItem();
            string aml = @"<AML>
                            <Item type='Workflow Process' action='get' >
	                            <Relationships>
	                              <Item type='Workflow Process Activity' action='get'>
	                                <related_id>{0}</related_id>
	                              </Item>
	                            </Relationships>
	
                            </Item>
                           </AML>";
            aml = string.Format(aml,Act.getID());
            WFP = inn.applyAML(aml);

            Item moreAppTemplate = inn.newItem();
            aml = @"<AML>
                            <Item type='Activity' action='get' >
	                          <id>BCC9CEA04A3543809A6262A26063BC62</id>
                                <Relationships>
                                   <Item type='Activity Method' action='get'>
	                               </Item>
                                </Relationships>
                            </Item>
                           </AML>";
            moreAppTemplate = inn.applyAML(aml);  //加签模板，可以在模板中设定需要出发的方法


            Item moreApp = moreAppTemplate.clone(true);
            moreApp.setProperty("name","More Approval");
            moreApp.setProperty("x",Act.getProperty("x"));
            moreApp.setProperty("y", (int.Parse(Act.getProperty("y"))-100).ToString());
            moreApp.apply();

            aml = @"<AML>                 
                        <Item type='Activity Assignment' action='add'>
                            <related_id>{0}</related_id>
                            <source_id>{1}</source_id>
                            <voting_weight>100</voting_weight>
	                    </Item>              
                    </AML>";
            aml = string.Format(aml, Assignment.getID(), moreApp.getID());

        //    MessageBox.Show(aml);

            Item AA = inn.newItem();
            AA = inn.applyAML(aml);
           

            Item WFPA = inn.newItem("Workflow Process Activity", "add");
            WFPA.setProperty("source_id",WFP.getID());
            WFPA.setProperty("related_id", moreApp.getID());
            WFPA.apply();


            Item tran1 = inn.newItem("Workflow Process Path", "add");
            tran1.setProperty("source_id", Act.getID());
            tran1.setProperty("related_id", moreApp.getID());
            tran1.setProperty("name","More Approve");
            tran1.setProperty("x","-80");
            tran1.setProperty("y", "-50");
            tran1.apply();

            int breakX = int.Parse(Act.getProperty("x")) + 35;
            int breakY = int.Parse(Act.getProperty("y")) - 50;
            string segments = breakX.ToString() + "," + breakY.ToString();

            Item tran2 = inn.newItem("Workflow Process Path", "add");
            tran2.setProperty("source_id", moreApp.getID());
            tran2.setProperty("related_id", Act.getID());
            tran2.setProperty("name", "Agree");
            tran2.setProperty("segments", segments);
            tran2.setProperty("x", "25");
            tran2.setProperty("y", "50");
            tran2.apply();


             breakX = int.Parse(Act.getProperty("x")) +90;
             breakY = int.Parse(Act.getProperty("y")) -50;
             segments = breakX.ToString() + "," + breakY.ToString();

            Item tran3 = inn.newItem("Workflow Process Path", "add");
            tran3.setProperty("source_id", moreApp.getID());
            tran3.setProperty("related_id", Act.getID());
            tran3.setProperty("name", "Disagree");
            tran3.setProperty("segments", segments);
            tran3.setProperty("x", "80");
           tran3.setProperty("y", "50");
            tran3.apply();

            //   MessageBox.Show("OK");

            return tran1;


        }
    }
}
