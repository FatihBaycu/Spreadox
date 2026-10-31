# Stage 1: Build & Publish
FROM mcr.microsoft.com/dotnet/sdk:8.0 AS build
WORKDIR /source

# Copy Directory.Build.props and project files for caching restore layer
COPY ["Directory.Build.props", "./"]
COPY ["src/Conq.Noble.Spreadox/Conq.Noble.Spreadox.csproj", "src/Conq.Noble.Spreadox/"]
COPY ["samples/ReportBuilderMvc/ReportBuilderMvc.csproj", "samples/ReportBuilderMvc/"]
RUN dotnet restore "samples/ReportBuilderMvc/ReportBuilderMvc.csproj"

# Copy full source (including README.md, LICENSE, CHANGELOG.md required by csproj None items)
COPY . .

# Publish in Release mode, disabling ContinuousIntegrationBuild to avoid missing git metadata errors
WORKDIR "/source/samples/ReportBuilderMvc"
RUN dotnet publish "ReportBuilderMvc.csproj" -c Release -o /app/publish /p:UseAppHost=false /p:ContinuousIntegrationBuild=false

# Stage 2: Runtime image
FROM mcr.microsoft.com/dotnet/aspnet:8.0 AS final
WORKDIR /app
COPY --from=build /app/publish .

ENV ASPNETCORE_ENVIRONMENT=Production
ENV PORT=8080
EXPOSE 8080

ENTRYPOINT ["dotnet", "ReportBuilderMvc.dll"]
