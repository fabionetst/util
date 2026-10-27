#include <iostream>
#include <iomanip>
#include <fstream>
#include <vector>
#include <sstream>
#include <algorithm>
#include <xlsxio_read.h> // Cabeçalho oficial de leitura da XLSX I/O

class Xls {

    private:

    std::string strNomeArq,strValor;

    xlsxioreader xlsLeitor;
    xlsxioreadersheet xlsFolha;

    bool bXlsLeitor;
    bool bXlsFolha;
    
    public:

    inline const std::string impCelula() const
    {
        return strValor;
    }
    inline bool okLeitor() const
    {
        return bXlsLeitor;
    }
    inline bool okFolha() const
    {
        return bXlsFolha;
    }
    inline bool lerLinha()
    {
        return (xlsxioread_sheet_next_row(xlsFolha))? true: false;
    }
    
    bool lerCelula()
    {
        char *charValor;

        if((charValor = xlsxioread_sheet_next_cell(xlsFolha)) == NULL)
            return false;

        strValor = charValor;
        xlsxioread_free(charValor);
        
        return true;
    }

    Xls(const std::string strNomeArq)
    {
        bXlsFolha = bXlsLeitor = false;
        this-> strNomeArq = strNomeArq;

        if((xlsLeitor = xlsxioread_open(strNomeArq.c_str())) != NULL)
            bXlsLeitor = true;
        if (bXlsLeitor and (xlsFolha = xlsxioread_sheet_open(xlsLeitor, NULL, 
                  XLSXIOREAD_SKIP_EMPTY_ROWS)) != NULL)
            bXlsFolha = true;
    }
    ~Xls()
    {
        if(bXlsFolha)
            xlsxioread_sheet_close(xlsFolha);
        if(bXlsLeitor)
            xlsxioread_close(xlsLeitor);
    }
};

class FUtil {
    
    
    public:

    bool estatisticas(int nVezes=0)
    {
        std::vector<std::string> vLinha;
        std::vector<std::vector<std::string>> vTabela;
        std::vector<int> vnQuant(26,0);
        std::vector<float> vfPorCento(26,0);
        std::vector<std::pair<int,int>> vpQuantInd;
        std::vector<int> viLotofacil;

        std::fstream arq;

        unsigned int nTotalLin,nTotalCol;

        int ultimaLin;
        
        Xls lf ("../db/lotofacil.xlsx");

        if(!lf.okLeitor()) {

            std::cout<< "Erro ao tentar abrir arquivo.\n";
            return false;
        }
        else if(!lf.okFolha()) {

            std::cout<< "Erro ao tentar abrir tabela.\n";
            return false;
        }
        else if(!lf.lerLinha()) {

            std::cout<< "Tabela vazia.\n";
            return false;
        }

        while(lf.lerLinha()) {

            while(lf.lerCelula()) {

                if(lf.impCelula() == "")
                    vLinha.push_back(" ");
                else 
                    vLinha.push_back(lf.impCelula());
            }

            vTabela.push_back(vLinha);
            vLinha.clear();
        }

        arq.open("../db/lotofacil.tdb",std::fstream::out|std::fstream::trunc);
        
        for(auto &t: vTabela) {

            for(auto &l: t)
                arq<< l<< '#';

            arq<< std::endl;
        }
        arq.close();

        nTotalLin = vTabela.size();
        nTotalCol = vTabela[0].size();
        ultimaLin = nTotalLin -1;

        for(int nC = 2; nC < 17; nC++)
            viLotofacil.push_back(std::stoi(vTabela[ultimaLin][nC]));

        if(nVezes <= 0)
            nVezes = nTotalLin;

        for(int iL = (nTotalLin -nVezes); iL < (ultimaLin -1); iL++)
            for(int iC = 2; iC < 17; iC++)
                vnQuant[std::stoi(vTabela[iL][iC])]++;
        for(int nN = 1; nN <= 25; nN++)
            vpQuantInd.push_back({vnQuant[nN],nN});
        std::sort(vpQuantInd.rbegin(),vpQuantInd.rend());
        

        arq.open("../db/quantlf.tdb",std::fstream::out|std::fstream::trunc);
        for(auto &par: vpQuantInd) {

            vfPorCento[par.second] = ((float) par.first /nVezes *100);
            arq<< std::setw(2)<< std::setfill('0')<< par.second<< ';'
               << std::setw(2)<< std::setfill('0')<< par.first<< ';'
               << std::fixed<< std::setprecision(2)<< vfPorCento[par.second];

            if(std::find(viLotofacil.begin(),viLotofacil.end(),par.second) != viLotofacil.end())
                arq<< ";*";
            
            arq<< std::endl;
        }
        arq.close();

        arq.open("../db/gquantlf.tdb",std::fstream::out|std::fstream::trunc);
        arq<< "\t\t\tMédia: "<< nVezes<< std::endl;
        for(int nC = 0; nC < 25; nC++) {

            arq<< std::setw(2)<< std::setfill('0')<< vpQuantInd[nC].second<< " "
               << std::string(vfPorCento[vpQuantInd[nC].second] -1,'*');

            if(std::find(viLotofacil.begin(),viLotofacil.end(),vpQuantInd[nC].second) != viLotofacil.end())
                arq<< 'x';
            else
                arq<< '*';

            arq<< " - "<< vfPorCento[vpQuantInd[nC].second]<< "% "<< std::endl;
        }

        arq<< "\nN. acertos 15: "<< std::setfill(' ')
           << std::setw(6)<< std::right<< vTabela[ultimaLin][17]<< " "
           << vTabela[ultimaLin][18]<< " "
           << vTabela[ultimaLin][19]<< "\n"

           << "\nN. acertos 14: "
           << std::setw(6)<< std::right<< vTabela[ultimaLin][20]<< " "
           << std::setw(10)<< std::right<< vTabela[ultimaLin][21]

           << "\nN. acertos 13: "
           << std::setw(6)<< std::right<< vTabela[ultimaLin][22]<< " "
           << std::setw(10)<< std::right<< vTabela[ultimaLin][23]

           << "\nN. acertos 12: "
           << std::setw(6)<< std::right<< vTabela[ultimaLin][24]<< " "
           << std::setw(10)<< std::right<< vTabela[ultimaLin][25]

           << "\nN. acertos 11: "
           << std::setw(6)<< std::right<< vTabela[ultimaLin][26]<< " "
           << std::setw(10)<< std::right<< vTabela[ultimaLin][27];


        arq.close();
        
        return true;
    }

    

    FUtil()
    {

    }

} util;

int main()
{
    util.estatisticas(30);
    
    return 0;
}