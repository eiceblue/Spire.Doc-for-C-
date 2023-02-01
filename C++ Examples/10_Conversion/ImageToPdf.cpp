#include "pch.h"
using namespace Spire::Doc;

int main() {
	wstring input_path = DATAPATH;
	wstring inputFile = input_path + L"Image.png";
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"ImageToPdf.pdf";

	//Create a new document
	Document* doc = new Document();
	//Create a new section
	Section* section = doc->AddSection();
	//Create a new paragraph
	Paragraph* paragraph = section->AddParagraph();
	//Add a picture for paragraph
	DocPicture* picture = paragraph->AppendPicture(inputFile.c_str());
	//Set the page size to the same size as picture
	section->GetPageSetup()->SetPageSize(new SizeF(picture->GetWidth(), picture->GetHeight()));
	//Set A4 page size
	section->GetPageSetup()->SetPageSize(PageSize::A4());
	//Set the page margins
	section->GetPageSetup()->GetMargins()->SetTop(10.0f);
	section->GetPageSetup()->GetMargins()->SetLeft(25.0f);

	doc->SaveToFile(outputFile.c_str(), FileFormat::PDF);
	doc->Close();

}
