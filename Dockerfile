FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS base
WORKDIR /app
EXPOSE 10000
ENV ASPNETCORE_URLS=http://0.0.0.0:10000

FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
ARG BUILD_CONFIGURATION=Release
WORKDIR /src
COPY ["TranscriptApp/TranscriptApp.csproj", "TranscriptApp/"]
RUN dotnet restore "./TranscriptApp/TranscriptApp.csproj"
COPY . .
WORKDIR "/src/TranscriptApp"
RUN dotnet build "./TranscriptApp.csproj" -c $BUILD_CONFIGURATION -o /app/build

FROM build AS publish
ARG BUILD_CONFIGURATION=Release
RUN dotnet publish "./TranscriptApp.csproj" -c $BUILD_CONFIGURATION -o /app/publish /p:UseAppHost=false

FROM base AS final
WORKDIR /app
COPY --from=publish /app/publish .
ENTRYPOINT ["dotnet", "TranscriptApp.dll"]