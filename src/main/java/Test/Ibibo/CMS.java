package main.java.Test.Ibibo;
import java.util.Scanner;

public class CMS {

	public static void main(String[] args) throws Exception {
		
		Scanner in = new Scanner(System.in);
		
		String abc;
		boolean flag=false;
		
		System.out.println("Enter name to find number from excel sheet");
		abc = in.nextLine();
		
		ReadDataFromExcel xls =new ReadDataFromExcel("C:\\Users\\Jitender.Ahuja\\Desktop\\CMS.xlsx");
		int abc1 = xls.getRowCount("Sheet1");
		
		for (int i=2; i<=abc1;i++)
		{
			
			String data = xls.getCellData("Sheet1", 0, i);
			if (data.equalsIgnoreCase(abc) )
			{
			
				flag=true;
				String Number = xls.getCellData("Sheet1", 1, i);
				System.out.println("Number of " +data+ " is: "+Number);
			
			}
			
			
		} if(flag==false)
		
		{
			System.out.println("Not Exist");
		}		
			
			in.close();
	}

}