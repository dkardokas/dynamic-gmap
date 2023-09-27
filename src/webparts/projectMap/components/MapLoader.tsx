/* eslint-disable @typescript-eslint/no-unused-vars */
/* eslint-disable @typescript-eslint/explicit-function-return-type */
import * as React from "react";
import { useEffect } from "react";
import { MarkerClusterer } from "@googlemaps/markerclusterer";
import { IItems } from "@pnp/sp/items";

interface IMapLoaderProps {
    spListItems: IItems;
    startLat: number;
    startLon: number;
}

export default function MapLoader(props: IMapLoaderProps) {

    async function initMap(): Promise<void> {

        await google.maps.importLibrary("maps") as google.maps.MapsLibrary;
        await google.maps.importLibrary("marker") as google.maps.MarkerLibrary;
        const itemArray: any[] = await props.spListItems();
        
        const map = new google.maps.Map(
            document.getElementById("dg_map") as HTMLElement,
            {
                zoom: 4,
                center: { lat: props.startLat, lng: props.startLon },
                mapId: 'DG_MAP_ID',
            }
        );


        const infoWindow = new google.maps.InfoWindow({
            content: "",
            disableAutoPan: true,
        });

        // Add some markers to the map.


        const markers = itemArray.filter(spItem => {
            const latValue = parseFloat(spItem.GeoLoc.Latitude);
            const lonValue = parseFloat(spItem.GeoLoc.Longitude);
            return !isNaN(latValue) && !isNaN(lonValue);
        }).map((spItem, i) => {
            const label = spItem.Title;

            const latValue = parseFloat(spItem.GeoLoc.Latitude);
            const lonValue = parseFloat(spItem.GeoLoc.Longitude);
            const position = { lat: latValue, lng: lonValue };


            // const glyphImg = document.createElement('img');
            // glyphImg.src = 'https://icons.iconarchive.com/icons/icons8/ios7/24/Industry-Wind-Turbine-icon.png';

            // const glyphSvgPinElement = new google.maps.marker.PinElement({
            //     glyph: glyphImg,
            // });

            const marker = new google.maps.marker.AdvancedMarkerElement({
                position,
                //content: glyphSvgPinElement.element,
            });

            // markers can only be keyboard focusable when they have click listeners
            // open info window when marker is clicked
            marker.addListener("click", () => {
                infoWindow.setContent(label);
                infoWindow.open(map, marker);
            });
            return marker;
        });

        // Add a marker clusterer to manage the markers.
        // eslint-disable-next-line no-new
        new MarkerClusterer({ markers, map });
    }


    useEffect(() => {
        // Update the document title using the browser API
        initMap().catch((e) => { console.log(e) });
    });

    return (
        <div id="dg_map" style={{ width: "100%", height: "35rem" }}>
            Error: Google Map API failed to load.
        </div>
    )
}