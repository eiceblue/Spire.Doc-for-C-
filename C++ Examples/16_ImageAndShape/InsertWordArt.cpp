#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path  + L"InsertWordArt.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertWordArt.docx";

	//Create Word document.
	Document* doc = new Document();

	//Load Word document.
	doc->LoadFromFile(inputFile.c_str());

	//Add a paragraph.
	Paragraph* paragraph = doc->GetSections()->GetItem(0)->AddParagraph();

	//Add a shape.
	ShapeObject* shape = paragraph->AppendShape(250, 70, ShapeType::TextWave4);

	//Set the position of the shape.
	shape->SetVerticalPosition(20);
	shape->SetHorizontalPosition(80);

	//set the text of WordArt.
	shape->GetWordArt()->SetText(L"Thanks for reading.");

	//Set the fill color.
	shape->SetFillColor(Color::GetRed());

	//Set the border color of the text.
	shape->SetStrokeColor(Color::GetYellow());

	//Save docx file.
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}
