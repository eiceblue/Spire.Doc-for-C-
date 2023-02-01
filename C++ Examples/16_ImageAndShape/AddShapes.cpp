#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddShapes.docx";

	//Create Word document.
	Document* doc = new Document();
	Section* sec = doc->AddSection();
	Paragraph* para = sec->AddParagraph();
	int x = 60, y = 40, lineCount = 0;
	for (int i = 1; i < 20; i++)
	{
		if (lineCount > 0 && lineCount % 8 == 0)
		{
			para->AppendBreak(BreakType::PageBreak);
			x = 60;
			y = 40;
			lineCount = 0;
		}
		//Add shape and set its size and position.
		ShapeObject* shape = para->AppendShape(50, 50, (ShapeType)i);
		shape->SetHorizontalOrigin(HorizontalOrigin::Page);
		shape->SetHorizontalPosition(x);
		shape->SetVerticalOrigin(VerticalOrigin::Page);
		shape->SetVerticalPosition(y + 50);
		x = x + static_cast<int>(shape->GetWidth()) + 50;
		if (i > 0 && i % 5 == 0)
		{
			y = y + static_cast<int>(shape->GetHeight()) + 120;
			lineCount++;
			x = 60;
		}

	}
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	doc->Close();
	delete doc;
}