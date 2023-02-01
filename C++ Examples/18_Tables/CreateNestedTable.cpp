#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"CreateNestedTable.docx";

	//Create a new document
	Document* doc = new Document();
	Section* section = doc->AddSection();

	//Add a table
	Table* table = section->AddTable(true);
	table->ResetCells(2, 2);

	//Set column width
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->SetCellWidth(70.0F, CellWidthType::Point);
	table->GetRows()->GetItem(1)->GetCells()->GetItem(0)->SetCellWidth(70.0F, CellWidthType::Point);
	table->AutoFit(AutoFitBehaviorType::AutoFitToWindow);

	//Insert content to cells
	table->GetRows()->GetItem(0)->GetCells()->GetItem(0)->AddParagraph()->AppendText(L"Spire.Doc for C++");
	wstring text = L"Spire.Doc for C++ is a professional Word"
		L"C++ library specifically designed for developers to create,"
		L"read, write, convert and print Word document files from any C++"
		L"platform with fast and high quality performance.";
	table->GetRows()->GetItem(0)->GetCells()->GetItem(1)->AddParagraph()->AppendText(text.c_str());

	//Add a nested table to cell(first row, second column)
	Table* nestedTable = table->GetRows()->GetItem(0)->GetCells()->GetItem(1)->AddTable(true);
	nestedTable->ResetCells(4, 3);
	nestedTable->AutoFit(AutoFitBehaviorType::AutoFitToContents);

	//Add content to nested cells
	nestedTable->GetRows()->GetItem(0)->GetCells()->GetItem(0)->AddParagraph()->AppendText(L"NO.");
	nestedTable->GetRows()->GetItem(0)->GetCells()->GetItem(1)->AddParagraph()->AppendText(L"Item");
	nestedTable->GetRows()->GetItem(0)->GetCells()->GetItem(2)->AddParagraph()->AppendText(L"Price");

	nestedTable->GetRows()->GetItem(1)->GetCells()->GetItem(0)->AddParagraph()->AppendText(L"1");
	nestedTable->GetRows()->GetItem(1)->GetCells()->GetItem(1)->AddParagraph()->AppendText(L"Pro Edition");
	nestedTable->GetRows()->GetItem(1)->GetCells()->GetItem(2)->AddParagraph()->AppendText(L"$799");

	nestedTable->GetRows()->GetItem(2)->GetCells()->GetItem(0)->AddParagraph()->AppendText(L"2");
	nestedTable->GetRows()->GetItem(2)->GetCells()->GetItem(1)->AddParagraph()->AppendText(L"Standard Edition");
	nestedTable->GetRows()->GetItem(2)->GetCells()->GetItem(2)->AddParagraph()->AppendText(L"$599");

	nestedTable->GetRows()->GetItem(3)->GetCells()->GetItem(0)->AddParagraph()->AppendText(L"3");
	nestedTable->GetRows()->GetItem(3)->GetCells()->GetItem(1)->AddParagraph()->AppendText(L"Free Edition");
	nestedTable->GetRows()->GetItem(3)->GetCells()->GetItem(2)->AddParagraph()->AppendText(L"$0");

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
