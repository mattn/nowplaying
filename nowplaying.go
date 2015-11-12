package main

import (
	"fmt"
	"log"
	"os"
	"path/filepath"
	"time"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
	"github.com/mattn/go-gntp"
)

func main() {
	cwd, _ := os.Getwd()
	artworkjpg := filepath.Join(cwd, "artwork.jpg")

	gc := gntp.NewClient()
	gc.AppName = "nowplaying"
	gc.Register([]gntp.Notification{{"default", "", true}})

	ole.CoInitialize(0)
	unk, err := oleutil.CreateObject("iTunes.Application")
	if err != nil {
		log.Fatal(err)
	}
	unk.AddRef()
	dsp := unk.MustQueryInterface(ole.IID_IDispatch)

	prev := ""
	for {
		func() {
			defer func() {
				recover()
			}()
			track := oleutil.MustGetProperty(dsp, "CurrentTrack").ToIDispatch()
			if track != nil {
				name := oleutil.MustGetProperty(track, "Name").ToString()
				artist := oleutil.MustGetProperty(track, "artist").ToString()
				album := oleutil.MustGetProperty(track, "Album").ToString()
				curr := fmt.Sprintf("%s/%s\n%s", name, artist, album)
				if curr != prev {
					prev = curr
					artwork := oleutil.MustGetProperty(track, "Artwork").ToIDispatch()
					item := oleutil.MustGetProperty(artwork, "Item", 1).ToIDispatch()
					icon := ""
					if item != nil {
						_, err = oleutil.CallMethod(item, "SaveArtworkToFile", artworkjpg)
						if err != nil {
							log.Print(err)
						} else {
							icon = artworkjpg
						}
					}
					log.Print(curr)
					err = gc.Notify(&gntp.Message{
						Event: "default",
						Title: "iTunes",
						Text:  curr,
						Icon:  icon,
					})
					if err != nil {
						log.Print(err)
					}
				}
			}
		}()
		time.Sleep(3 * time.Second)
	}
}
