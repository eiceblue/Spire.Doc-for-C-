#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"InsertOLE.docx";

	//create a document
	Document* doc = new Document();

	//add a section
	Section* sec = doc->AddSection();

	//add a paragraph
	Paragraph* par = sec->AddParagraph();

	//load the image
	DocPicture* picture = new DocPicture(doc);
	wstring imagePath = input_path + L"Excel.png";
	picture->LoadImageSpire(imagePath.c_str());

	//insert the OLE
	wstring filePath = input_path + L"example.xlsx";
	DocOleObject* obj = par->AppendOleObject(filePath.c_str(), picture, OleObjectType::ExcelWorksheet);
	doc->SaveToFile(outputFile.c_str(), FileFormat::Docx2013);
	doc->Close();
	delete doc;
}
