#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ImageWaterMark.docx";

	//Open a Word document as template.
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Insert the imgae watermark.
	InsertImageWatermark(document);
	//Save as docx file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;

}

void InsertImageWatermark(Document* document)
{
	wstring input_path = DATAPATH;
	PictureWatermark* picture = new PictureWatermark();
	wstring imagePath = input_path + L"ImageWatermark.png";
	picture->SetPicture(Image::FromFile(imagePath.c_str()));
	picture->SetScaling(250);
	picture->SetIsWashout(false);
	document->SetWatermark(picture);
}

