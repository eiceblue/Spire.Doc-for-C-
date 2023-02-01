#include "pch.h"
using namespace Spire::Doc;


int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ConvertedTemplate.docx";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ToImage.png";

	//Create word document
	Document* document = new Document();
	document->LoadFromFile(inputFile.c_str());

	//Save doc file.
	Stream* imageStream = document->SaveToImages(0, ImageFormat::GetPng());
	imageStream->Save(outputFile.c_str());
	document->Close();
	delete document;
	imageStream->Dispose();
}
