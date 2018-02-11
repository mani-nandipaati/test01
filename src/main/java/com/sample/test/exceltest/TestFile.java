package com.sample.test.exceltest;

import java.io.File;
import java.io.IOException;

public class TestFile {
	public static void main(String[] args) {
		File file = new File("I:\\spring_interview_question.PDF");
		try {
			System.out.println(file.getCanonicalFile());
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		System.out.println("test2....");
	}
	
	public static void test2() {
		System.out.println("test2....");
		System.out.println("test222....");
	}
	

}
