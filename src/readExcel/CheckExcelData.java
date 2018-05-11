package readExcel;

import object.ExcelObject;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;


public class CheckExcelData {

    public static void main(String[] args) {

        List<ExcelObject> sheetDataList = getSheetData("C:\\Users\\admin\\Desktop\\config\\zss.xlsx", "raheem");
        System.out.println("表中有" + sheetDataList.size() + "行数据。");

        //1、检查Name（index=1）列中字段长度是否小于等于30
//        checkLength(sheetDataList);

        //2、连词小写
//        checkPrepLowwerCase(sheetDataList);

        //3、医学专业名词
//        checkMedWords(sheetDataList);

        //4、检查缩写
        checkAbbr(sheetDataList);
    }

    public static void checkAbbr(List<ExcelObject> sheetDataList) {
        for (int rowNo = 0; rowNo < sheetDataList.size(); rowNo++) {
            ExcelObject obj = sheetDataList.get(rowNo);
            String name = obj.getName();
            String desc = obj.getDesc();

            //入院
            if (desc.contains("入院")) {
                if (!name.contains("Admission") && !name.contains("Admsn")) {
                    System.out.println("医学专业名词使用有误：入院。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }

            //Doctor	Dct
            checkAbbrSingle(obj, "Doctor", "Dct");
            //Medical 	Med
            checkAbbrSingle(obj, "Medical", "Med");
            //Department	Dept
            checkAbbrSingle(obj, "Department", "Dept");
            //Operation	Oper
            checkAbbrSingle(obj, "Operation", "Oper");
            //Assistant	Assist
            checkAbbrSingle(obj, "Assistant", "Assist");
            //Address	Addr
            checkAbbrSingle(obj, "Address", "Addr");
            //Examination	Exam
            checkAbbrSingle(obj, "Examination", "Exam");
            //Descriptin	Desc
            checkAbbrSingle(obj, "Descriptin", "Desc");
            //广义的中医	CN_Med
            if (desc.contains("中医")) {
                if (!name.contains("CN_Med")) {
                    System.out.println(String.format("名词缩写有误：%s正确的应为：%s。%s%s%s%s%s","中医","CN_Med",obj.getRowNum(), obj.getCode(), obj.getName(), obj.getDesc(), obj.getComeFrom()));
                }
            }
            //中药   	TCM
            if (desc.contains("中药")) {
                if (!name.contains("TCM")) {
                    System.out.println(String.format("名词缩写有误：%s正确的应为：%s。%s%s%s%s%s","中药","TCM", obj.getRowNum(), obj.getCode(), obj.getName(), obj.getDesc(), obj.getComeFrom()));
                }
            }
            //医保	Insur
            if (desc.contains("医保")) {
                if (!name.contains("Insur") && name.contains("Insurance")) {
                    System.out.println(String.format("名词缩写有误：%s正确的应为：%s。%s%s%s%s%s","医保,Insurance","Insur", obj.getRowNum(), obj.getCode(), obj.getName(), obj.getDesc(), obj.getComeFrom()));
                }
            }
            //检验	Lab
            if (desc.contains("检验")||desc.contains("检查")) {
                if (!name.contains("Lab")&&!name.contains("Exam")) {
                    System.out.println(String.format("名词缩写有误：%s正确的应为：%s。%s%s%s%s%s","检验、检查","Lab、Exam", obj.getRowNum(), obj.getCode(), obj.getName(), obj.getDesc(), obj.getComeFrom()));
                }
            }
            //Organization	Org
            checkAbbrSingle(obj, "Organization", "Org");
            //Register	Reg
            checkAbbrSingle(obj, "Register", "Reg");
            //Prescription	Rx
            checkAbbrSingle(obj, "Prescription", "Rx");
            //Statistics	Stat
            checkAbbrSingle(obj, "Statistics", "Stat");
            //Present	Pres
            checkAbbrSingle(obj, "Present", "Pres");
            //Telephone	Tel
            checkAbbrSingle(obj, "Telephone", "Tel");
            //Hypertension	Hyper
            checkAbbrSingle(obj, "Hypertension", "Hyper");
            //Exercise	Exer
            checkAbbrSingle(obj, "Exercise", "Exer");
            //Frequency
            checkAbbrSingle(obj, "Frequency", "Freq");
            //Cigarette(烟)	Ciga
            checkAbbrSingle(obj, "Cigarette", "Ciga");

            continue;
        }
    }

    public static void checkAbbrSingle(ExcelObject obj, String errorStr, String corrStr) {
        String name = obj.getName();
        if (name.contains(errorStr)) {
            System.out.println(String.format("名词缩写有误：%s正确的应为：%s。%s%s%s%s%s", errorStr, corrStr, obj.getRowNum(), obj.getCode(), obj.getName(), obj.getDesc(), obj.getComeFrom()));
        }
    }

    //医学专业名词http://wiki.iflytek.com/pages/viewpage.action?pageId=186975029
    public static void checkMedWords(List<ExcelObject> sheetDataList) {
        for (int rowNo = 0; rowNo < sheetDataList.size(); rowNo++) {
            ExcelObject obj = sheetDataList.get(rowNo);
            String name = obj.getName();
            String desc = obj.getDesc();

            //入院
            if (desc.contains("入院")) {
                if (!name.contains("Admission") && !name.contains("Admsn")) {
                    System.out.println("医学专业名词使用有误：入院。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //出院
            if (desc.contains("出院")) {
                if (!name.contains("Discharge")) {
                    System.out.println("医学专业名词使用有误：出院。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }

            //门诊 no 患者
            if (desc.contains("门诊") && !desc.contains("患者")) {
                if (!name.contains("Outpat") || name.contains("Outpatient")) {
                    System.out.println("医学专业名词使用有误：门诊。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //住院 no 患者
            if (desc.contains("住院") && !desc.contains("患者")) {
                if (!name.contains("Inpat") || name.contains("Inpatient")) {
                    System.out.println("医学专业名词使用有误：住院。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //住院患者
            if (desc.contains("住院患者")) {
                if (!name.contains("Inpatient")) {
                    System.out.println("医学专业名词使用有误：住院患者。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //转诊/转出
            if (desc.contains("转诊") || desc.contains("转出")) {
                if (!name.contains("Transfer")) {
                    System.out.println("医学专业名词使用有误：转诊/转出。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }

            //接诊/ 转入	Receive
            if (desc.contains("接诊") || desc.contains("转入")) {
                if (!name.contains("Receive")) {
                    System.out.println("医学专业名词使用有误：接诊/转入。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //诊断	Diag
            if (desc.contains("诊断")) {
                if (!name.contains("Diag")) {
                    System.out.println("医学专业名词使用有误：诊断。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //药品/药物	Drug
            if (desc.contains("药品") || desc.contains("药物")) {
                if (!name.contains("Drug")) {
                    System.out.println("医学专业名词使用有误：药品/药物" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //医嘱	Order
            if (desc.contains("医嘱")) {
                if (!name.contains("Order")) {
                    System.out.println("医学专业名词使用有误：医嘱。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //报告	Rept/Report
            if (desc.contains("报告")) {
                if (!name.contains("Rept") && !name.contains("Report")) {
                    System.out.println("医学专业名词使用有误：报告。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //临床路径	CPW  (Clinical pathway)
            if (desc.contains("临床路径")) {
                if (!name.contains("CPW")) {
                    System.out.println("医学专业名词使用有误：临床路径。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //病案号	Case_No
            if (desc.contains("病案号")) {
                if (!name.contains("Case_No")) {
                    System.out.println("医学专业名词使用有误：病案号。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //门诊号	Outpat_No
            if (desc.contains("门诊号")) {
                if (!name.contains("Outpat_No")) {
                    System.out.println("医学专业名词使用有误：门诊号。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //住院号	Inpat_No
            if (desc.contains("住院号")) {
                if (!name.contains("Inpat_No")) {
                    System.out.println("医学专业名词使用有误：住院号。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //床位号	Bed_No
            if (desc.contains("床位号")) {
                if (!name.contains("Bed_No")) {
                    System.out.println("医学专业名词使用有误：床位号。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //中药	TCM
            if (desc.contains("中药")) {
                if (!name.contains("TCM")) {
                    System.out.println("医学专业名词使用有误：中药。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //草药	Herb
            if (desc.contains("草药")) {
                if (!name.contains("Herb")) {
                    System.out.println("医学专业名词使用有误：草药。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //西药	Medicine
            if (desc.contains("西药")) {
                if (!name.contains("Medicine")) {
                    System.out.println("医学专业名词使用有误：西药。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //检验	Lab
            if (desc.contains("检验")) {
                if (!name.contains("Lab")) {
                    System.out.println("医学专业名词使用有误：检验。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //检查、广义的检验检查	Exam
            if (desc.contains("检查") || desc.contains("检验")) {
                if (!name.contains("Exam")) {
                    System.out.println("医学专业名词使用有误：检查、广义的检验检查。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //手术/操作	Surgery
            if (desc.contains("手术") && !desc.contains("手术者")) {
                if (!name.contains("Surgery")) {
                    System.out.println("医学专业名词使用有误：手术。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //手术者	Surgery
            if (desc.contains("手术者")) {
                if (!name.contains("Surgeon")) {
                    System.out.println("医学专业名词使用有误：手术者。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //操作 no手术	Oper/Operation
            if (desc.contains("操作") && !desc.contains("手术")) {
                if (!name.contains("Oper") && !name.contains("Operation")) {
                    System.out.println("医学专业名词使用有误：操作。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //医保	Insur/Insurance
            if (desc.contains("医保")) {
                if (!name.contains("Insur") && !name.contains("Insurance")) {
                    System.out.println("医学专业名词使用有误：医保。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            // 申请单号	 Apply_No
            if (desc.contains("申请单号")) {
                if (!name.contains("Apply_No")) {
                    System.out.println("医学专业名词使用有误：申请单号。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            // 项目   	 Item
            if (desc.contains("项目")) {
                if (!name.contains("Item")) {
                    System.out.println("医学专业名词使用有误：项目。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //处方	Rx
            if (desc.contains("处方")) {
                if (!name.contains("Rx")) {
                    System.out.println("医学专业名词使用有误：处方。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //注册	Reg
            if (desc.contains("注册")) {
                if (!name.contains("Reg")) {
                    System.out.println("医学专业名词使用有误：注册。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //科室	Dept
            if (desc.contains("科室")) {
                if (!name.contains("Dept")) {
                    System.out.println("医学专业名词使用有误：科室。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //执业	Cert
            if (desc.contains("执业")) {
                if (!name.contains("Cert")) {
                    System.out.println("医学专业名词使用有误：执业。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //统计	Stat
            if (desc.contains("统计")) {
                if (!name.contains("Stat")) {
                    System.out.println("医学专业名词使用有误：统计。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //标本	Specimen
            if (desc.contains("标本")) {
                if (!name.contains("Specimen")) {
                    System.out.println("医学专业名词使用有误：标本。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //疾病	Disease
            if (desc.contains("疾病")) {
                if (!name.contains("Disease")) {
                    System.out.println("医学专业名词使用有误：疾病。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //中医	CN_Med
            if (desc.contains("中医")) {
                if (!name.contains("CN_Med")) {
                    System.out.println("医学专业名词使用有误：中医。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //西医	Med
            if (desc.contains("西医")) {
                if (!name.contains("Med_")) {
                    System.out.println("医学专业名词使用有误：西医。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //症状/症候	Symptom
            if (desc.contains("症状") || desc.contains("症候")) {
                if (!name.contains("Symptom")) {
                    System.out.println("医学专业名词使用有误：症状/症候。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //证候（中医的专业术语）	Syndrome
            if (desc.contains("证候")) {
                if (!name.contains("Syndrome")) {
                    System.out.println("医学专业名词使用有误：证候。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //术前	PreOper
            if (desc.contains("术前")) {
                if (!name.contains("PreOper")) {
                    System.out.println("医学专业名词使用有误：术前。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //术后	PostOper
            if (desc.contains("术后")) {
                if (!name.contains("PostOper")) {
                    System.out.println("医学专业名词使用有误：术后。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //初步	Init
            if (desc.contains("初步")) {
                if (!name.contains("Init")) {
                    System.out.println("医学专业名词使用有误：初步。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //修正	Revise
            if (desc.contains("修正")) {
                if (!name.contains("Revise")) {
                    System.out.println("医学专业名词使用有误：修正。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //确定	Final
            if (desc.contains("确定")) {
                if (!name.contains("Final")) {
                    System.out.println("医学专业名词使用有误：确定。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //补充	Supply
            if (desc.contains("补充")) {
                if (!name.contains("Supply")) {
                    System.out.println("医学专业名词使用有误：补充。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //费用	Fee
            if (desc.contains("费用")) {
                if (!name.contains("Fee")) {
                    System.out.println("医学专业名词使用有误：费用。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //类别/类型	只有一级类别：    Type
            if (desc.contains("类别") || desc.contains("类型")) {
                if (!name.contains("Type") && !name.contains("Cat")) {
                    System.out.println("医学专业名词使用有误：类别/类型。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }

            //体检	Physical_Exam
            if (desc.contains("体检")) {
                if (!name.contains("Physical_Exam")) {
                    System.out.println("医学专业名词使用有误：体检。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //检查	Exam
            if (desc.contains("检查")) {
                if (!name.contains("Exam")) {
                    System.out.println("医学专业名词使用有误：检查。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }


            //主治医生/主治医师	Attending_Dct
            if (desc.contains("主治医生") || desc.contains("主治医师")) {
                if (!name.contains("Attending_Dct")) {
                    System.out.println("医学专业名词使用有误：主治医生/主治医师。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }

            //频率/频次	Freq
            if (desc.contains("频率") || desc.contains("频次")) {
                if (!name.contains("Freq")) {
                    System.out.println("医学专业名词使用有误：频率/频次。" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }


            continue;
        }
    }

    ///连词小写（and/or/per）
    public static void checkPrepLowwerCase(List<ExcelObject> sheetDataList) {
        for (int rowNo = 0; rowNo < sheetDataList.size(); rowNo++) {
            ExcelObject obj = sheetDataList.get(rowNo);
            String name = obj.getName();

            //and
            if (name.toLowerCase().contains("_and_")) {
                if (!name.contains("_and_")) {
                    System.out.println("连词未小写：" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //or
            if (name.toLowerCase().contains("_or_")) {
                if (!name.contains("_or_")) {
                    System.out.println("连词未小写：" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }
            //per
            if (name.toLowerCase().contains("_per_")) {
                if (!name.contains("_per_")) {
                    System.out.println("连词未小写：" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
                }
            }

            continue;
        }
    }

    //检查Name（index=1）列中字段长度是否小于等于30
    public static void checkLength(List<ExcelObject> sheetDataList) {
        for (int rowNo = 0; rowNo < sheetDataList.size(); rowNo++) {
            ExcelObject obj = sheetDataList.get(rowNo);
            String name = obj.getName();
            if (name.length() > 30) {
                System.out.println("字段长度超过30：" + obj.getRowNum() + obj.getCode() + obj.getName() + obj.getDesc() + obj.getComeFrom());
            } else {
                //System.out.println("字段长度正常：" + obj.getRowNum());
            }
            continue;
        }
    }

    //获取表中数据，存储至list集合中
    public static List<ExcelObject> getSheetData(String excelPath, String item) {
        boolean isFound = false;
        File f = new File(excelPath);
        Workbook book = null;
        Sheet currSheet = null;
        int rowCount = 0;
        int colCount = 0;

        try {
            book = new XSSFWorkbook(new FileInputStream(f));
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }

        outerLoop:

        //取总行数
        try {
            currSheet = book.getSheetAt(0);
            rowCount = currSheet.getPhysicalNumberOfRows();
            colCount = currSheet.getRow(0).getPhysicalNumberOfCells();
        } catch (Exception e) {
            e.printStackTrace();
        }

        //存取表数据至集合
        List<ExcelObject> sheetDataList = new ArrayList<ExcelObject>();
        ExcelObject tempObj = null;
        for (int rowNo = 0; rowNo < rowCount; rowNo++) {
            Row currRow = currSheet.getRow(rowNo);
            tempObj = new ExcelObject();
            tempObj.setRowNum(rowNo);
            tempObj.setCode(currRow.getCell(0).getStringCellValue());
            tempObj.setName(currRow.getCell(1).getStringCellValue());
            tempObj.setDesc(currRow.getCell(2).getStringCellValue());
            tempObj.setComments(currRow.getCell(3).getStringCellValue());
//            tempObj.setComeFrom(currRow.getCell(4).getStringCellValue());
            tempObj.setComeFrom("sb.");
            sheetDataList.add(tempObj);
        }
        return sheetDataList;

    }

}
