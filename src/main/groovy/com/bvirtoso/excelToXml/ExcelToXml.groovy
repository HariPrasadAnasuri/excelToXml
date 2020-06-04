package com.bvirtoso.excelToXml

import com.bvirtoso.excelToXml.service.PrepareXml
import org.springframework.boot.CommandLineRunner
import org.springframework.boot.SpringApplication
import org.springframework.boot.autoconfigure.SpringBootApplication

@SpringBootApplication
class ExcelToXml implements CommandLineRunner{

	final PrepareXml prepareXml

	public ExcelToXml(PrepareXml prepareXml1){
		this.prepareXml = prepareXml1
	}

	static void main(String[] args) {
		SpringApplication.run(ExcelToXml, args)
	}

	@Override
	void run(String... args) throws Exception {
		prepareXml.prepareXmlFromExcel("./excelFiles/personAddress.xlsx", "./xmlFiles/personsInXmlFormat.xml")
	}
}
