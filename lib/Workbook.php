<?php

namespace Excel;

final class Workbook extends Writer\Workbook
{
    const GRIGIO_SCURO = 60;
    const GRIGIO_MEDIO = 61;
    const GRIGIO_CHIARO = 62;

    private $righePerPagina = 60000;

    private $identita;

    private $formati;

    public function __construct($filename)
    {
        parent::__construct($filename);

        $this->setCustomColor(self::GRIGIO_SCURO,   hexdec('7f'), hexdec('7f'), hexdec('7f'));
        $this->setCustomColor(self::GRIGIO_MEDIO,   hexdec('cc'), hexdec('cc'), hexdec('cc'));
        $this->setCustomColor(self::GRIGIO_CHIARO,  hexdec('e8'), hexdec('e8'), hexdec('e8'));

        $this->identita = new StileCella\Testo();
    }

    public function setRighePerPagina($righePerPagina)
    {
        $this->righePerPagina = (int) $righePerPagina;
    }

    public function scriviTabella(Tabella $tabella)
    {
        $this->scriviIntestazioneTabella($tabella);
        $tabelle = array($tabella);

        $count = 0;
        $rigaIntestazioni = true;
        foreach ($tabella->getDati() as $riga) {
            ++$count;

            if ($tabella->getRigaCorrente() >= $this->righePerPagina) {
                $tabella = $tabella->dividiTabellaSuNuovoSheet($this->addWorksheet(uniqid()));
                $tabelle[] = $tabella;
                $this->scriviIntestazioneTabella($tabella);
                $rigaIntestazioni = true;
            }

            if ($rigaIntestazioni) {
                $this->scriviIntestazioneColonne($tabella, $riga);

                $rigaIntestazioni = false;
            }

            $this->scriviRiga($tabella, $riga);
        }

        if (count($tabelle) > 1) {
            $tabella = reset($tabelle);
            $firstSheet = $tabella->getActiveSheet();
            // Il massimo di caratteri per il nome di un foglio Excel e' 30
            $nomeOriginale = substr($firstSheet->name, 0, 21);

            $contatoreFogli = 0;
            $totaleFogli = count($tabelle);
            foreach ($tabelle as $tabella) {
                ++$contatoreFogli;
                $tabella->getActiveSheet()->name = sprintf('%s (%s|%s)', $nomeOriginale, $contatoreFogli, $totaleFogli);
            }
        }

        if ($tabella->getBloccaRiquadri()) {
            foreach ($tabelle as $tabella) {
                $tabella->getActiveSheet()->freezePanes(array($tabella->getRigaIniziale() + 2, 0));
            }
        }

        if ($count === 0) {
            $tabella->incrementaRiga();
            $tabella->getActiveSheet()->writeString($tabella->getRigaCorrente(), $tabella->getColonnaCorrente(), 'Nessun dato per questa estrazione');
            $tabella->incrementaRiga();
        }

        $tabella->setCount($count);

        return end($tabelle);
    }

    private function scriviIntestazioneTabella(Tabella $tabella)
    {
        $tabella->ripristinaColonna();
        $tabella->getActiveSheet()->writeString($tabella->getRigaCorrente(), $tabella->getColonnaCorrente(), $tabella->getIntestazione());
        $tabella->incrementaRiga();
    }

    private function scriviIntestazioneColonne(Tabella $tabella, array $riga)
    {
        $colonnaCollection = $tabella->getColonnaCollection();
        $chiaviColonne = array_keys($riga);
        $this->generaFormati($chiaviColonne, $colonnaCollection);

        $tabella->ripristinaColonna();
        $titoli = array();
        foreach ($chiaviColonne as $titolo) {
            $larghezza = 10;
            $nuovoTitolo = ucwords(str_replace('_', ' ', $titolo));

            if (isset($colonnaCollection) and isset($colonnaCollection[$titolo])) {
                $larghezza = $colonnaCollection[$titolo]->getLarghezza();
                $nuovoTitolo = $colonnaCollection[$titolo]->getIntestazione();
            }

            $tabella->getActiveSheet()->setColumn($tabella->getColonnaCorrente(), $tabella->getColonnaCorrente(), $larghezza);
            $titoli[$titolo] = $nuovoTitolo;

            $tabella->incrementaColonna();
        }

        $this->scriviRiga($tabella, $titoli, 'titolo');
    }

    private function scriviRiga(Tabella $tabella, array $riga, $tipo = null)
    {
        $tabella->ripristinaColonna();

        foreach ($riga as $chiave => $contenuto) {
            $stileCella = $this->identita;
            $formato = null;
            if (isset($this->formati[$chiave])) {
                if ($tipo === null) {
                    $tipo = (($tabella->getRigaCorrente() % 2)
                        ? 'zebra_scura'
                        : 'zebra_chiara'
                    );
                }
                $stileCella = $this->formati[$chiave]['stile_cella'];
                $formato = $this->formati[$chiave][$tipo];
            }

            $write = 'write';
            if (get_class($stileCella) === get_class($this->identita)) {
                $write = 'writeString';
            }

            $contenuto = $stileCella->decorateValue($contenuto);
            $contenuto = $this->sanitize($contenuto);

            $tabella->getActiveSheet()->{$write}($tabella->getRigaCorrente(), $tabella->getColonnaCorrente(), $contenuto, $formato);

            $tabella->incrementaColonna();
        }

        $tabella->incrementaRiga();
    }

    private function sanitize($value)
    {
        static $sanitizeMap;

        if ($sanitizeMap === null) {
            $sanitizeMap = array(
                '&amp;'     => '&',
                '&lt;'      => '<',
                '&gt;'      => '>',
                '&apos;'    => "'",
                '&quot;'    => '"',
            );
        }

        $value = str_replace(
            array_keys($sanitizeMap),
            array_values($sanitizeMap),
            $value
        );
        $value = mb_convert_encoding($value, 'Windows-1252');

        return $value;
    }

    private function generaFormati(array $titoli, ColonnaCollection $colonnaCollection = null)
    {
        $this->formati = array();
        foreach ($titoli as $chiave) {
            $header = $this->addFormat();
            $header->setColor('black');
            $header->setSize(8);
            $header->setBold();
            $header->setFgColor(self::GRIGIO_MEDIO);
            $header->setTextWrap();
            $header->setAlign('center');

            $rigaChiara = $this->addFormat();
            $rigaChiara->setColor('black');
            $rigaChiara->setSize(8);
            $rigaChiara->setFgColor('white');
            $rigaChiara->SetBorderColor(self::GRIGIO_SCURO);

            $rigaScura = $this->addFormat();
            $rigaScura->setColor('black');
            $rigaScura->setSize(8);
            $rigaScura->setFgColor(self::GRIGIO_CHIARO);
            $rigaScura->SetBorderColor(self::GRIGIO_SCURO);

            $this->formati[$chiave] = array(
                'stile_cella'   => null,
                'titolo'        => $header,
                'zebra_chiara'  => $rigaChiara,
                'zebra_scura'   => $rigaScura,
            );

            $stileCella = $this->identita;
            if (isset($colonnaCollection) and isset($colonnaCollection[$chiave])) {
                $stileCella = $colonnaCollection[$chiave]->getStileCella();
            }

            $stileCella->styleCell($rigaChiara);
            $stileCella->styleCell($rigaScura);

            $this->formati[$chiave]['stile_cella'] = $stileCella;
        }
    }
}
