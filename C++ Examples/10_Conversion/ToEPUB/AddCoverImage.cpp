#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"ToEpub.doc";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"AddCoverImage.epub";

	Document* doc = new Document();
	doc->LoadFromFile(inputFile.c_str());
	DocPicture* picture = new DocPicture(doc);
	picture->LoadImageSpire(Image::FromFile((input_path + L"Cover.png").c_str()));
	doc->SaveToEpub(outputFile.c_str(), picture);
	doc->Close();
	delete doc;
}
