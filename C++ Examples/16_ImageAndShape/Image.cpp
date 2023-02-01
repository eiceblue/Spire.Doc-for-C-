#include "pch.h"
using namespace Spire::Doc;
using namespace Spire::Common;

int main() {
	wstring output_path = OUTPUTPATH;
	wstring outputFile = output_path + L"Image.docx";


	//Create a document
	Document* document = new Document();

	//Add a seciton
	Section* section = document->AddSection();

	//insert image
	InsertImage(section);

	//Save doc file.
	document->SaveToFile(outputFile.c_str(), FileFormat::Docx);
	document->Close();
	delete document;
}

void InsertImage(Section* section) {
	//Add paragraph
	Paragraph* paragraph = section->AddParagraph();
	paragraph->GetFormat()->SetHorizontalAlignment(HorizontalAlignment::Left);

	wstring input_path = DATAPATH;
	wstring imagePath = input_path + L"Spire.Doc.png";
	Spire::Common::Image* ima = Image::FromFile(imagePath.c_str());

	//Add a image and set its width and height
	DocPicture* picture = paragraph->AppendPicture(ima);
	picture->SetWidth(100);
	picture->SetHeight(100);

	paragraph = section->AddParagraph();
	paragraph->GetFormat()->SetLineSpacing(20.0f);
	TextRange* tr = paragraph->AppendText(L"Spire.Doc for C++ is a professional Word C++ library specially designed for developers to create, read, write, convert and print Word document files from any C++( C#, VBCPP, ASPCPP) platform with fast and high quality performance. ");
	tr->GetCharacterFormat()->SetFontName(L"Arial");
	tr->GetCharacterFormat()->SetFontSize(14);

	section->AddParagraph();
	paragraph = section->AddParagraph();
	paragraph->GetFormat()->SetLineSpacing(20.0f);
	tr = paragraph->AppendText(L"As an independent Word C++ component, Spire.Doc for C++ doesn't need Microsoft Word to be installed on the machine. However, it can incorporate Microsoft Word document creation capabilities into any developers' C++ applications.");
	tr->GetCharacterFormat()->SetFontName(L"Arial");
	tr->GetCharacterFormat()->SetFontSize(14);
}
