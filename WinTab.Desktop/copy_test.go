package main

import (
	"reflect"
	"strings"
	"testing"
)

func TestPunctuationFreeCopy(t *testing.T) {
	for _, entry := range []struct {
		name string
		copy copyCatalog
	}{{"zh-CN", zhCopy}, {"en-US", enCopy}} {
		val := reflect.ValueOf(entry.copy)
		typ := val.Type()
		for i := range val.NumField() {
			text := val.Field(i).String()
			if text == "" {
				t.Fatalf("%s.%s is empty", entry.name, typ.Field(i).Name)
			}
			for _, suffix := range []string{".", "。", "...", "…"} {
				if strings.HasSuffix(text, suffix) {
					t.Errorf("%s.%s ends with forbidden punctuation %q: %q", entry.name, typ.Field(i).Name, suffix, text)
				}
			}
		}
	}
}

func TestEmbeddedBrandAssetsAndIcons(t *testing.T) {
	if len(logoPNG) == 0 || logoBitmap == nil {
		t.Fatal("logo PNG was not embedded or decoded")
	}
	if iconFolderOpen == nil || iconHistory == nil || iconTuning2 == nil || iconInfoCircle == nil ||
		iconAltArrowRight == nil || iconSun == nil || iconMoon == nil || iconTranslate == nil ||
		iconCheck == nil || iconDangerCircle == nil || iconRefresh == nil || iconKeyboard == nil {
		t.Fatal("one or more embedded SVG icons failed to parse")
	}
}
