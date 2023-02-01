#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Template_HtmlFile1.html";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"HtmlToImage.png";

	//Create Word document.
	Document* document = new Document();

	//Load the file from disk.
	document->LoadFromFile(inputFile.c_str(), FileFormat::Html, XHTMLValidationType::None);

	//Save to image. You can convert HTML to BMP, JPEG, PNG, GIF, Tiff，etc.
	Stream* imageStream = document->SaveToImages(0, ImageFormat::GetPng());
	imageStream->Save(outputFile.c_str());
	document->Close();
	delete document;
	imageStream->Dispose();
}
