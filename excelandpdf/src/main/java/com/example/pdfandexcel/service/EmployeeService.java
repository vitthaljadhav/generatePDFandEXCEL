package com.example.pdfandexcel.service;

import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import com.example.pdfandexcel.model.Employee;

public interface EmployeeService {

	List<Employee> allEmployee();

	boolean createpdf(List<Employee> employees, ServletContext context, HttpServletRequest request,
			HttpServletResponse response);

	boolean createExcel(List<Employee> employees, ServletContext context, HttpServletRequest request,
			HttpServletResponse response);

}
