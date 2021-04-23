package com.ops.reportgenerator.entitites;

import java.io.FileInputStream;

import org.springframework.web.multipart.MultipartFile;

public class GetDoc {
private String date;
private String name;
private int counter;
private MultipartFile file;
public String getDate() {
	return date;
}
public void setDate(String date) {
	this.date = date;
}
public String getName() {
	return name;
}
public void setName(String name) {
	this.name = name;
}
public int getCounter() {
	return counter;
}
public void setCounter(int counter) {
	this.counter = counter;
}
public MultipartFile getFile() {
	return file;
}
public void setFile(MultipartFile file) {
	this.file = file;
}


 
}
