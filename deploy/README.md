For details see this [article](https://htmlacademy.ru/blog/articles/bot-hosting)

# Steps to deploy the app
VPS hosting: [reg.ru](https://www.reg.ru)

## Add public SSH key to VPS
hosting --> VPS --> Settings --> SSH keys

## Install all necessary software
hosting --> VPS --> Select my server --> Open console

Check operation system: `cat /etc/os-release`
- for Ubuntu/Debian run `apt -y install nodejs npm screen`
- for CentOS run `yum -y install nodejs npm screen`

## Get host IP address
hosting --> VPS --> Home --> IP address

## Copy files to the remote VPS server
Connect VPS with SFTP using [FileZilla](https://filezilla-project.org)
- Host: `sftp://host_ip_address`
- User: `root`
- Password: see email `Создан сервер Amethyst Neon`
- Port: leave empty

Copy project files in the directory `./home/my_app_directory`

## Create a service for auto restart the app
Create file `my_app_name.service` using template bellow and put it here `./lib/systemd/system`

```
  [Unit]
  Description=my_app_description
  After=network.target

  [Service]
  ExecStart=npm run start
  ExecReload=npm run start
  WorkingDirectory=/home/my_app_directory/
  KillMode=process
  Restart=always
  RestartSec=5

  [Install]
  WantedBy=multi-user.target
```

## Install app's dependencies
hosting --> VPS --> Select my server --> Open console
- run `cd ./home/my_app_directory`
- run `cd npm ci`

## Set up apache web server
### SSL sertifiates
Folder `/etc/ssl`

File `fullchain.crt` - сертификат +  промежуточный сертификат + корневой сертификат
```
  -----BEGIN CERTIFICATE-----
  ...xxx
  -----END CERTIFICATE-----
  -----BEGIN CERTIFICATE-----
  ...xxx
  -----END CERTIFICATE-----
  -----BEGIN CERTIFICATE-----
  ...xxx
  -----END CERTIFICATE-----
```

File `private.key` - приватный ключ

```
  -----BEGIN RSA PRIVATE KEY-----
  ...xxx
  -----END RSA PRIVATE KEY-----
```

## Run the app
hosting --> VPS --> Select my server --> Open console
- run `screen` (allows to close console and keep app running)
- Press space
- run `cd ./home/my_app_directory`
- run `npm run service:enable`
- run `npm run service:start`

## Check app status
hosting --> VPS --> Select my server --> Open console
- run `cd ./home/my_app_directory`
- run `npm run service:status`

## Restart the app
hosting --> VPS --> Select my server --> Open console
- run `cd ./home/my_app_directory`
- run `npm run service:restart`

## Stop the app
hosting --> VPS --> Select my server --> Open console
- run `cd ./home/my_app_directory`
- run `npm run service:stop`