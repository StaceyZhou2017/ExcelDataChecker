package object;

public class ExcelObject {
    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    private int rowNum;
    private String Code;
    private String Name;

    public String getComeFrom() {
        return ComeFrom;
    }

    public void setComeFrom(String comeFrom) {
        ComeFrom = comeFrom;
    }

    private String ComeFrom;

    public String getCode() {
        return Code;
    }

    public void setCode(String code) {
        Code = code;
    }

    public String getName() {
        return Name;
    }

    public void setName(String name) {
        Name = name;
    }

    public String getDesc() {
        return Desc;
    }

    public void setDesc(String desc) {
        Desc = desc;
    }

    public String getComments() {
        return Comments;
    }

    public void setComments(String comments) {
        Comments = comments;
    }

    private String Desc;
    private String Comments;


}
