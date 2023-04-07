#ifndef CSVWRITER_H
#define CSVWRITER_H
#include <fstream>
#include <iostream>
#include <sstream>
#include <typeinfo>
#include <codecvt>

class CSVWriter
{
    public:
        CSVWriter(){
            this->firstRow = true;
            this->seperator = L";";
            this->columnNum = -1;
            this->valueCount = 0;
        }

        CSVWriter(int numberOfColums){
            this->firstRow = true;
            this->seperator = L";";
            this->columnNum = numberOfColums;
            this->valueCount = 0;
        }

        CSVWriter(std::wstring seperator){
            this->firstRow = true;
            this->seperator = seperator;
            this->columnNum = -1;
            this->valueCount = 0;
        }

        CSVWriter(std::wstring seperator, int numberOfColums){
            this->firstRow = true;
            this->seperator = seperator;
            this->columnNum = numberOfColums;
            this->valueCount = 0;
            std::wcout << this->seperator << std::endl;
        }

        CSVWriter& add(const wchar_t *str){
            return this->add(std::wstring(str));
        }

        CSVWriter& add(wchar_t *str){
            return this->add(std::wstring(str));
        }

		CSVWriter& add(_bstr_t str){
            return this->add(std::wstring(str));
        }

        CSVWriter& add(std::wstring str){
            //if " character was found, escape it
            size_t position = str.find(L"\"",0);
            bool foundQuotationMarks = position != std::wstring::npos;
            while(position != std::wstring::npos){
                str.insert(position,L"\"");
                position = str.find(L"\"",position + 2);
            }
            if(foundQuotationMarks){
                str = L"\"" + str + L"\"";
            }else if(str.find(this->seperator) != std::wstring::npos){
                //if seperator was found and string was not escapted before, surround string with "
                str = L"\"" + str + L"\"";
            }
            return this->add<std::wstring>(str);
        }

        template<typename T>
        CSVWriter& add(T str){
            if(this->columnNum > -1){
                //if autoNewRow is enabled, check if we need a line break
                if(this->valueCount == this->columnNum ){
                    this->newRow();
                }
            }
            if(valueCount > 0)
                this->ss << this->seperator;
            this->ss << str;
            this->valueCount++;

            return *this;
        }

        template<typename T>
        CSVWriter& operator<<(const T& t){
            return this->add(t);
        }

        void operator+=(CSVWriter &csv){
            this->ss << std::endl << csv;
        }

        std::wstring toString(){
            return ss.str();
        }

        friend std::wostream& operator<<(std::wostream& os, CSVWriter & csv){
            return os << csv.toString();
        }

        CSVWriter& newRow(){
            if(!this->firstRow || this->columnNum > -1){
                ss << std::endl;
            }else{
                //if the row is the first row, do not insert a new line
                this->firstRow = false;
            }
            valueCount = 0;
            return *this;
        }

        bool writeToFile(std::wstring filename){
			return writeToFile(filename.c_str(),false);
        }

		bool writeToFile(std::wstring filename, bool append){
			return writeToFile(filename.c_str(), append);
		}

		bool writeToFile(const wchar_t *filename){
			return writeToFile(filename,false);
		}

        bool writeToFile(const wchar_t *filename, bool append){
            if (!append) {
                // write BOM
                std::ofstream file;
                file.open(filename, std::ios::out | std::ios::trunc | std::ios::binary);
                if (!file.is_open())
                    return false;
				file << (unsigned char)0xEF << (unsigned char)0xBB << (unsigned char)0xBF;
				file.close();
			}
            std::wofstream file;
            file.imbue(std::locale(std::locale(), new std::codecvt_utf8<wchar_t, 0x10ffff>));
            file.open(filename, std::ios::out | std::ios::app | std::ios::binary );
            if(!file.is_open())
                return false;
            if(append)
                file << std::endl;
            file << this->toString();
            file.close();
            return file.good();
        }

        void enableAutoNewRow(int numberOfColumns){
            this->columnNum = numberOfColumns;
        }

        void disableAutoNewRow(){
            this->columnNum = -1;
        }
    protected:
        bool firstRow;
        std::wstring seperator;
        int columnNum;
        int valueCount;
        std::wstringstream ss;

};

#endif // CSVWRITER_H
