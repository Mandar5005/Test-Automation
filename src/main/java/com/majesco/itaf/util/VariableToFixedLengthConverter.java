package com.majesco.itaf.util;
import java.io.*;
import java.nio.file.Files;
public class VariableToFixedLengthConverter {

	public static int countFields(String s, String fieldSeperator) {
		
		int count=0;
		int index = fieldSeperator.length()-1;
		for(int i=0;i<s.length();i++) {
			if(s.charAt(i)==fieldSeperator.charAt(index)) {
				count++;
			}
		}
		return  count+1;
	}
	
	public static void createCopy(String source) {
		String destination = source.substring(0, source.indexOf('.')) + "_Response" + source.substring(source.indexOf('.'), source.length());
		try {
			File src = new File(source);
			File dest = new File(destination);
			Files.copy(src.toPath(), dest.toPath());
		}catch(Exception e) {
			e.printStackTrace();
			System.out.println(e);
		}
	}
	public static boolean convertToFixedLength(String fileName, String fieldSeperator) {
		boolean isEmpty=true;
		
		System.out.println("Saving Original Feed File......");
		createCopy(fileName);
		
		System.out.println("Converting to Fixed length file: "+fileName);
		try {
			File file = new File(fileName);
			FileReader fr= new FileReader(file);
			BufferedReader br = new BufferedReader(fr);
			StringBuffer sb = new StringBuffer();
			String line;
			System.out.println("Reading File......!");
			while((line=br.readLine())!=null) {
				isEmpty=false;
				if(line.length()==0) {
					continue;
				}
				int totalFields = countFields(line, fieldSeperator);
				String []words = line.split(fieldSeperator);
				for(int i=0;i<words.length;i++) {
					String word = words[i];
					int len = word.length();
					int spaces = 60-len;
					for(int j=1;j<=spaces;j++) {
						word = " " + word;
					}
					words[i] = word;
				}
				String newLine="";
				for(int i=0;i<words.length;i++) {
					
					if(i==0) {
						newLine=words[i];
					}
					else {
						newLine=newLine + "," + words[i];
					}
				}
				
				int numberOfFieldsToAdd = totalFields - words.length;
//				System.out.println("Number of fields to add: "+numberOfFieldsToAdd);
				for(int i=1;i<=numberOfFieldsToAdd;i++) {
					String word = ",";
					for(int j=1;j<=60;j++) {
						word = word + " ";
					}
					newLine = newLine + word;
				}
				sb.append(newLine);
				sb.append("\n");
			}
			System.out.println("Writing......");
			BufferedWriter bwr = new BufferedWriter(new FileWriter(new File(fileName)));
			bwr.write(sb.toString());
			bwr.flush();
			bwr.close();
			fr.close();
			System.out.println("Conversion to fixed length feed file completed.");
			
		}catch(IOException e) {
			e.printStackTrace();
			System.out.println(e);
			
		}
		return isEmpty;
	}
}