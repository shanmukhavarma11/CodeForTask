package example;

public class TestingData  extends InsertDataFromExcelToDb{
	static String dataInsert(String id1,String name) {
		return id1+name;
	}
public static void main(String[] args) {
	TestingData test=new TestingData();
	test.age="Shanmukha";
	test.firstname="lastname";
	System.out.println(test.age+" "+test.firstname+" "+" "+dataInsert(test.age,test.firstname));
}
}
