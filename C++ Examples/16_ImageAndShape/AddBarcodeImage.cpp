#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"SampleB_2.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddBarcodeImage.docx";

	//Open a Word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	wstring imgPath = output_path + L"barcode.png";

	//Add barcode image
	DocPicture* picture = document->GetSections()->GetItem(0)->AddParagraph()->AppendPicture(Image::FromFile(imgPath.c_str()));

	//Save to file
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}
