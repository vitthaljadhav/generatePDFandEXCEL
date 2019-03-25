package com.example.pdfandexcel.controlles;

import java.io.File;
import java.io.FileInputStream;
import java.io.OutputStream;
import java.util.Date;
import java.util.List;

import javax.servlet.ServletContext;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;

import com.example.pdfandexcel.model.Employee;
import com.example.pdfandexcel.service.EmployeeService;

@Controller
public class PdfExcelController {

	@Autowired
	private EmployeeService employeeService;
	@Autowired
	private ServletContext context;

	@RequestMapping(value = "/")
	public String allEmployee(Model model) {
		List<Employee> employees = employeeService.allEmployee();
		model.addAttribute("employees", employees);
		return "view/employees";
	}

	@RequestMapping(value = "/createdpdf")
	public void createpdf(HttpServletRequest request, HttpServletResponse response) {
		List<Employee> employees = employeeService.allEmployee();
		boolean isflag = employeeService.createpdf(employees, context, request, response);
		if (isflag) {
			String fullPath = request.getServletContext().getRealPath("/resources/reports/" + "employee" + ".pdf");
			System.out.println("Full Path :" + fullPath);

			fileDownload(fullPath, request, response, "employee.pdf");
		}
	}
	@RequestMapping(value = "/createdexcel")
	public void createexcel(HttpServletRequest request, HttpServletResponse response) {
		List<Employee> employees = employeeService.allEmployee();
		boolean isflag = employeeService.createExcel(employees, context, request, response);
		if (isflag) {
			String fullPath = request.getServletContext().getRealPath("/resources/reports/" + "employee" + ".xls");
			System.out.println("Full Path :" + fullPath);
			fileDownload(fullPath, request, response, "employee.xls");
		}
	}
	private void fileDownload(String fullPath, HttpServletRequest request, HttpServletResponse response,
			String fileName) {
		File file = new File(fullPath);
		final int BUFFER_SIZE = 4098;
		if (file.exists()) {
			try {
				FileInputStream inputStream = new FileInputStream(file);
				String mimeType = context.getMimeType(fullPath);
				response.setContentType(mimeType);
				response.setHeader("content-disposition", "attachment;filename=" + fileName);
				OutputStream outputStream = response.getOutputStream();

				byte[] buffer = new byte[BUFFER_SIZE];
				int bytesRead = -1;
				while ((bytesRead = inputStream.read(buffer)) != -1) {
					outputStream.write(buffer, 0, bytesRead);
				}
				inputStream.close();
				outputStream.close();
				file.delete();
			} catch (Exception e) {
				e.printStackTrace();
			}
		}

	}
}
