# Nest Thermostat Logger

Project to log data from Google Nest thermostats into a Google Sheet, using the Smart Device Management (SDM) API and Apps Script. It also uses the weather.gov API to add local weather data.

The SDM API was launched in September 2020 ([read more](https://developers.googleblog.com/2020/09/google-nest-device-access-console.html)).

# Setup

Follow the steps in this [Get Started](https://developers.google.com/nest/device-access/get-started) guide from Google to access your device(s).

**Note:** there is a one-time, non-refundable $5 charge to access the API.

The project must be created in the same Google account as the Nest home devices are connected to.

You will need to create a project in your [Google Cloud console](https://console.cloud.google.com/), where you can get OAuth ID and secret.

This project uses the [OAuth2 Apps Script library](https://github.com/googleworkspace/apps-script-oauth2).

# Useful API Documentation

[Get Device list](https://developers.google.com/nest/device-access/reference/rest/v1/enterprises.devices/list)

[Get Device data](https://developers.google.com/nest/device-access/reference/rest/v1/enterprises.devices/get)

And yes, it's possible to set your thermostat temperature from your Google Sheet, at this [endpoint](https://developers.google.com/nest/device-access/traits/device/thermostat-temperature-setpoint)

# More Info

Tutorial coming soon on [benlcollins.com](https://www.benlcollins.com/)
