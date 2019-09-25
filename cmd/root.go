package cmd

import (
	"encoding/csv"
	"fmt"
	"os"
	"path/filepath"
	"strings"

	"github.com/loadoff/excl"
	"github.com/pakuyuya/gopathmatch"
	"github.com/spf13/cobra"
	"golang.org/x/text/encoding/japanese"
	"golang.org/x/text/transform"
)

func Execute() {
	if err := rootCmd.Execute(); err != nil {
		fmt.Println(err)
		os.Exit(1)
	}
}

type Options struct {
	Encode string
	Csv    string
	Xlsx   string
}

var (
	o = &Options{}
)

func init() {
	rootCmd.Flags().StringVarP(&o.Encode, "encode", "e", "utf8", "CSVファイルのエンコード。utf8、sjisが指定可")
	rootCmd.Flags().StringVarP(&o.Csv, "csv", "c", "", "CSVファイル。ワイルド―カード指定可能")
	rootCmd.MarkFlagRequired("csv")
	rootCmd.Flags().StringVarP(&o.Xlsx, "xlsx", "x", "", "出力あとのExcelファイル。CSVファイルをワイルドカードで指定している場合は無視される")
}

var rootCmd = &cobra.Command{
	Use:   "csv2xlsx",
	Short: "CSVファイルをxlsxファイルに変換するツールです",
	Run: func(cmd *cobra.Command, args []string) {
		if strings.Index(o.Csv, "*") < 0 {
			dest := o.Xlsx
			if dest == "" {
				fname := filepath.Base(o.Csv)
				dest = fname[0:len(fname)-len(filepath.Ext(fname))] + ".xlsx"
			}
			convert(o.Csv, dest)
		} else {
			for _, fpath := range gopathmatch.Listup(o.Csv, gopathmatch.FlgFileOnly) {
				fname := filepath.Base(fpath)
				dest := fname[0:len(fname)-len(filepath.Ext(fname))] + ".xlsx"
				convert(fpath, dest)
			}
		}
	},
}

func convert(src string, dest string) {
	file, err := os.Open(src)
	if err != nil {
		fmt.Println(err.Error())
		return
	}
	defer file.Close()

	var csvReader *csv.Reader
	switch o.Encode {
	case "utf8":
		csvReader = csv.NewReader(file)
	case "sjis":
		csvReader = csv.NewReader(transform.NewReader(file, japanese.ShiftJIS.NewDecoder()))
	}

	var line []string
	w, _ := excl.Create()
	s, _ := w.OpenSheet("Sheet1")

	rNum := 1
	for {
		line, err = csvReader.Read()
		if err != nil && err.Error() == "EOF" {
			break
		}

		r := s.GetRow(rNum)
		for c, v := range line {
			r.GetCell(c + 1).SetString(v)
		}

		rNum = rNum + 1
	}

	s.Close()
	err = w.Save(dest)
	if err != nil {
		fmt.Println(err.Error())
	}
}
