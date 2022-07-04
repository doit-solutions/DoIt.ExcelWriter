FROM mcr.microsoft.com/dotnet/sdk:6.0.300-bullseye-slim
RUN apt update && apt install -y zsh git
ENV SHELL /bin/zsh
ADD https://github.com/JanDeDobbeleer/oh-my-posh3/releases/latest/download/posh-linux-amd64 /usr/local/bin/oh-my-posh
RUN chmod +x /usr/local/bin/oh-my-posh
ADD https://github.com/JanDeDobbeleer/oh-my-posh3/raw/main/themes/paradox.omp.json /root/downloadedtheme.json
RUN echo eval "$(oh-my-posh prompt init zsh --config /root/downloadedtheme.json)" >> /root/.zshrc
RUN dotnet tool install --global dotnet-outdated-tool
ENV PATH ${PATH}:/root/.dotnet/tools
