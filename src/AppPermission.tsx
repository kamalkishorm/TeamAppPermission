import React from "react";
import * as microsoftTeams from "@microsoft/teams-js";
import { initializeIcons } from "@uifabric/icons";
import { Icon } from "@fluentui/react/lib/Icon";

class AppPermission extends React.Component<any, any> {
  constructor(props: any) {
    super(props);
    this.state = {
      teamsContext: null,
      positionString: null,
      coords: {
        lat: 0,
        lng: 0,
      },
    };
    microsoftTeams.initialize();
    initializeIcons();
    this.getTeamsContext();
    this.getDevicePermission();
  }

  getTeamsContext = () => {
    microsoftTeams.getContext((context) => {
      this.setState({
        teamsContext: context,
      });
    });
  };

  getDevicePermission = () => {
    this.getLocation();
  };

  showLocation = (devicePosition: any) => {
    var latitude = devicePosition.coords.latitude;
    var longitude = devicePosition.coords.longitude;
    this.setState({
      coords: {
        lat: latitude,
        lng: longitude,
      },
      positionString: "Latitude : " + latitude + " Longitude: " + longitude,
    });
    //alert("Latitude : " + latitude + " Longitude: " + longitude);
  };

  errorHandler = (err: any) => {
    if (err.code === 1) {
      alert("Error: Access is denied!");
    } else if (err.code === 2) {
      alert("Error: Position is unavailable!");
    }
  };

  getLocation = () => {
    if (navigator.geolocation) {
      // timeout at 60000 milliseconds (60 seconds)
      var options = { timeout: 60000 };
      navigator.geolocation.getCurrentPosition(
        this.showLocation,
        this.errorHandler,
        options
      );
    } else {
      alert("Sorry, browser does not support geolocation!");
    }
  };
  render() {
    return (
      <div>
        <h2>Hi,</h2>
        <h3>UPN == {this.state?.teamsContext?.upn}</h3>
        <h3>
          Location == <Icon iconName="CompassNW" className="ms-IconExample" />{" "}
          {this.state?.positionString}
        </h3>
      </div>
    );
  }
}
export default AppPermission;
