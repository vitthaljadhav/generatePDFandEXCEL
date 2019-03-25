package com.example.pdfandexcel.repos;

import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;

import com.example.pdfandexcel.model.Employee;
@Repository
public interface EmployeeRespository extends JpaRepository<Employee, Long> {

}
