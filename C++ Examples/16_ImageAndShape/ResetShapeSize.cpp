#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"Shapes.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ResetShapeSize.docx";

	//Load Document
	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());

	//Get the first section and the first paragraph that contains the shape
	Section* section = doc->GetSections()->GetItem(0);
	Paragraph* para = section->GetParagraphs()->GetItem(0);

	//Get the second shape and reset the width and height for the shape
	ShapeObject* shape = dynamic_cast<ShapeObject*>(para->GetChildObjects()->GetItem(1));
	shape->SetWidth(200);
	shape->SetHeight(200);

	//Save and launch document
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}
