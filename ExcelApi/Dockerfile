# Build stage
FROM mcr.microsoft.com/dotnet/sdk:10.0 AS build

WORKDIR /src

COPY . .

RUN dotnet restore
RUN dotnet publish -c Release -o /app/publish

# Runtime stage
FROM mcr.microsoft.com/dotnet/aspnet:10.0

WORKDIR /app

# Install fonts for ClosedXML auto-sizing and Excel operations
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    fontconfig \
    fonts-dejavu-core && \
    rm -rf /var/lib/apt/lists/*

ENV ASPNETCORE_URLS=http://+:5621

EXPOSE 5621

COPY --from=build /app/publish .

ENTRYPOINT ["dotnet", "ExcelApi.dll"]