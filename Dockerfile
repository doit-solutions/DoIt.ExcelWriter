FROM mcr.microsoft.com/dotnet/sdk:7.0.201-jammy
ARG TARGETARCH
RUN apt update && apt install -y zsh git
ENV SHELL /bin/zsh
RUN wget -O /usr/local/bin/oh-my-posh https://github.com/JanDeDobbeleer/oh-my-posh3/releases/latest/download/posh-linux-`echo ${TARGETARCH} | sed s/aarch64/arm64/ | sed s/x86_64/amd64/`
RUN chmod +x /usr/local/bin/oh-my-posh
ADD https://github.com/JanDeDobbeleer/oh-my-posh3/raw/main/themes/paradox.omp.json /root/downloadedtheme.json
RUN echo eval "$(oh-my-posh prompt init zsh --config /root/downloadedtheme.json)" >> /root/.zshrc
RUN dotnet tool install --global dotnet-outdated-tool
ENV PATH ${PATH}:/root/.dotnet/tools
