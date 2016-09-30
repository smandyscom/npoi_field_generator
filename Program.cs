using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.IO;
using System.Xml.Serialization;
using System.Xml;
using System.Reflection;
namespace CodeGenerator
{
    public class Program
    {
        const String fullFormat = "\t{0:G}\t\t\t\t\t\t\t\t=\t{1:G}\t\t'\t{2:G}\t{3:G}\n";        // [name] = [id]    ' [type]  [description]
        const String simplestFormat = "\t{0:G}\t\t\t\t\t\t\t\t=\t{1:G}\n";        // [name] = [id]    ' [type]  [description]
        const String settingFileName = @".\MappingSet.xml";

        /// <summary>
        /// 
        /// </summary>
        enum argsDefinition : int
        {
            SOURCE_FILE = 0,
            //DESTINATION_CSV  =1,
            DESTINATION_FILE = 1,
            SETTING_FILE =2,
            //IS_APPEND_MODE = 2,
            //ENUM_NAME   = 3,
            TOTAL_ARGS,
        };

        

        /// <summary>
        /// 
        /// </summary>
        public abstract class mappingSetBase
        {
            protected abstract string HeadString{get;}
            protected abstract ISheet InterestedSheet { get; }
            protected abstract string TailString { get; }

            protected abstract string generateBodyString(XSSFRow row);

            public virtual void initialize(XSSFWorkbook workbook)
            {
                this.cachedWorkbook = workbook;
            }
            protected XSSFWorkbook cachedWorkbook = null;

            /// <summary>
            /// Formula is stored as formula itself in .xlsx , counldn't directly output evaluted value without any rendering.
            /// </summary>
            /// <param name="cell"></param>
            /// <returns></returns>
            protected string cellFormator(ICell cell){
                switch (cell.CellType)
                {
                    case CellType.Blank:
                    case CellType.Boolean:
                    case CellType.Error:
                    case CellType.Numeric:
                    case CellType.String:
                    case CellType.Unknown:
                    default:
                        return cell.ToString();
                    case CellType.Formula:
                        //XSSFFormulaEvaluator fe = new XSSFFormulaEvaluator(this.cachedWorkbook);
                        //return fe.EvaluateInCell(cell).StringCellValue;
                        return cell.StringCellValue;
                }
            }
            /// <summary>
            /// 
            /// </summary>
            /// <param name="content"></param>
            public void generateContent(StringBuilder content)
            {
                content.Append(this.HeadString);
                foreach (XSSFRow row in this.InterestedSheet)
                {
                    Console.Write(string.Format("Processing Row {0} ...", row.RowNum));

                    try
                    {
                        if (!row.Cells.First().ToString().StartsWith("//")){
                            content.Append(generateBodyString(row));
                            Console.Write(string.Format("Processed {0} \n", row.RowNum));
                        }
                        else {
                            Console.Write(string.Format("Skip {0} \n", row.RowNum));
                        }
                    }
                    catch (Exception)
                    {
                        content.Append(string.Format("Error {0} \n", row.RowNum));
                    }
                    
                }
                content.Append(this.TailString);
            }
        }

        /// <summary>
        /// Type-1 setting method
        /// </summary>
        public class mappingSet1 :
            mappingSetBase
        {
            enum cellDefinition : int
            {
                NAME = 0,
                ID = 1,
                TYPE = 2,
                LENGTH = 3,
                UNIT = 4,
                DESCRIPTION = 5,
                TOTAL_NUM
            };

            public String headString = "Enum\t{0}\n";
            public String tailString = "End Enum\n";

            public String fullFormat = "\t{0:G}\t\t\t\t\t\t\t\t=\t{1:G}\t\t'\t{2:G}\t\n";

            public int nameColumn = (int)cellDefinition.NAME;
            public int idColumn = (int)cellDefinition.ID;
            public int comment = (int)cellDefinition.DESCRIPTION;
            public int sheetIndex = 0;

            protected override string HeadString
            {
                get { return headString; }
            }
            protected override string TailString
            {
                get { return tailString; }
            }
            protected override ISheet InterestedSheet
            {
                get { return this.cachedWorkbook.GetSheetAt(this.sheetIndex); }
            }

            protected override string generateBodyString(XSSFRow row)
            {
                return string.Format(this.fullFormat,
                    this.cellFormator(row.GetCell(nameColumn, MissingCellPolicy.CREATE_NULL_AS_BLANK)),
                    this.cellFormator(row.GetCell(idColumn, MissingCellPolicy.CREATE_NULL_AS_BLANK)),
                    this.cellFormator(row.GetCell(comment, MissingCellPolicy.CREATE_NULL_AS_BLANK)));
            }
        }

        public class mappingSet2:
            mappingSetBase
        {
            public string identifier = "$2";

            protected List<string> tags = new List<string>();

            protected override ISheet InterestedSheet
            {
                get {
                    foreach (ISheet item in this.cachedWorkbook)
                    {
                        if (item.SheetName.Contains(identifier))
                        {
                            // analyse first row to get all tags

                            return item;
                        }
                    }
                    return null;
                }
            }

            protected override string HeadString
            {
                get { throw new NotImplementedException(); }
            }

            protected override string TailString
            {
                get { throw new NotImplementedException(); }
            }

            protected override string generateBodyString(XSSFRow row)
            {
                throw new NotImplementedException();
            }
        }


        static int Main(string[] args)
        {
            try
            {
                //------------------
                //  Append mode only
                //------------------
                Console.WriteLine(String.Format("Header Generator {0} Started....",typeof(Program).Assembly.GetName().Version));


                if (args.Length < (int)argsDefinition.TOTAL_ARGS)
                {
                    args = new string[(int)argsDefinition.TOTAL_ARGS];
                    args[(int)argsDefinition.SOURCE_FILE] = @".\input.xlsx";
                    args[(int)argsDefinition.DESTINATION_FILE] = @".\output.vb";       // test code
                    //args[(int)argsDefinition.ENUM_NAME] = "testEnum";
                    args[(int)argsDefinition.SETTING_FILE] = @".\configuration.xml";
                }

                System.IO.FileStream source = new System.IO.FileStream(args[(int)argsDefinition.SOURCE_FILE]
                    , System.IO.FileMode.Open
                    , FileAccess.Read
                    , FileShare.ReadWrite);
                System.IO.StreamWriter destination = new System.IO.StreamWriter(args[(int)argsDefinition.DESTINATION_FILE]
                    , false
                    , Encoding.Default);

                //--------------------
                //  Check setting file
                //---------------------
                mappingSetBase setting = null;
                
                if (!File.Exists(args[(int)argsDefinition.SETTING_FILE]))
                {
                    // not existed , created new one
                    StreamWriter sw = new StreamWriter(args[(int)argsDefinition.SETTING_FILE]);
                    XmlSerializer xs = new XmlSerializer(typeof(mappingSet1));
                    xs.Serialize(sw, setting);
                    sw.Close();

                    Console.WriteLine("File Not Found Generating Template Mapping File Only");
                    return 0;
                }
                else
                {
                    XmlDocument xd = new XmlDocument();
                    xd.Load(args[(int)argsDefinition.SETTING_FILE]);
                    //decide serialization type according to root element 
                    Type deserializationType = Type.GetType(string.Format("{0}.Program+{1}",
                        "CodeGenerator",
                        xd.DocumentElement.Name));

                    // existed , open
                    StreamReader sr = new StreamReader(args[(int)argsDefinition.SETTING_FILE]);
                    XmlSerializer xs = new XmlSerializer(deserializationType);
                    setting = (mappingSetBase)xs.Deserialize(sr);
                    Console.WriteLine("Readin Setting :" + sr.ToString());
                    sr.Close();
                }


                //-----------
                //  Generating
                //------------
                Console.WriteLine("START GENERATING");

                StringBuilder content = new StringBuilder( );

                XSSFWorkbook sourceWorkBook = new XSSFWorkbook(source);
                setting.initialize(sourceWorkBook);
                //format head string,
                setting.generateContent(content);

                destination.Write(content);    //Hsien . 2014.10.17

                source.Close();
                destination.Close();

                Console.WriteLine("Generated");
                return 0;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                Console.WriteLine(e.StackTrace);
                throw;
            }//catch

 
        }//main
    }
}
