﻿<?xml version="1.0" encoding="utf-8"?>
<Project ToolsVersion="15.0" xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <Import Project="$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props" Condition="Exists('$(MSBuildExtensionsPath)\$(MSBuildToolsVersion)\Microsoft.Common.props')" />
  <PropertyGroup>
    <Configuration Condition=" '$(Configuration)' == '' ">Debug</Configuration>
    <Platform Condition=" '$(Platform)' == '' ">AnyCPU</Platform>
    <ProjectGuid>{8A790EF7-B1CC-4A47-B420-E54093E5874A}</ProjectGuid>
    <OutputType>WinExe</OutputType>
    <RootNamespace>ShiftReportApp1</RootNamespace>
    <AssemblyName>ShiftReportApp1</AssemblyName>
    <TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
    <FileAlignment>512</FileAlignment>
    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>
    <Deterministic>true</Deterministic>
    <TargetFrameworkProfile />
    <PublishUrl>publish\</PublishUrl>
    <Install>true</Install>
    <InstallFrom>Disk</InstallFrom>
    <UpdateEnabled>false</UpdateEnabled>
    <UpdateMode>Foreground</UpdateMode>
    <UpdateInterval>7</UpdateInterval>
    <UpdateIntervalUnits>Days</UpdateIntervalUnits>
    <UpdatePeriodically>false</UpdatePeriodically>
    <UpdateRequired>false</UpdateRequired>
    <MapFileExtensions>true</MapFileExtensions>
    <ApplicationRevision>0</ApplicationRevision>
    <ApplicationVersion>1.0.0.%2a</ApplicationVersion>
    <IsWebBootstrapper>false</IsWebBootstrapper>
    <UseApplicationTrust>false</UseApplicationTrust>
    <BootstrapperEnabled>true</BootstrapperEnabled>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Debug|AnyCPU' ">
    <PlatformTarget>AnyCPU</PlatformTarget>
    <DebugSymbols>true</DebugSymbols>
    <DebugType>full</DebugType>
    <Optimize>false</Optimize>
    <OutputPath>bin\Debug\</OutputPath>
    <DefineConstants>DEBUG;TRACE</DefineConstants>
    <ErrorReport>prompt</ErrorReport>
    <WarningLevel>4</WarningLevel>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Release|AnyCPU' ">
    <PlatformTarget>AnyCPU</PlatformTarget>
    <DebugType>pdbonly</DebugType>
    <Optimize>true</Optimize>
    <OutputPath>bin\Release\</OutputPath>
    <DefineConstants>TRACE</DefineConstants>
    <ErrorReport>prompt</ErrorReport>
    <WarningLevel>4</WarningLevel>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="ClosedXML, Version=0.102.1.0, Culture=neutral, PublicKeyToken=fd1eb21b62ae805b, processorArchitecture=MSIL">
      <HintPath>packages\ClosedXML.0.102.1\lib\netstandard2.0\ClosedXML.dll</HintPath>
    </Reference>
    <Reference Include="DocumentFormat.OpenXml, Version=2.16.0.0, Culture=neutral, PublicKeyToken=8fb06cb64d019a17, processorArchitecture=MSIL">
      <HintPath>packages\DocumentFormat.OpenXml.2.16.0\lib\net46\DocumentFormat.OpenXml.dll</HintPath>
    </Reference>
    <Reference Include="ExcelNumberFormat, Version=1.1.0.0, Culture=neutral, PublicKeyToken=23c6f5d73be07eca, processorArchitecture=MSIL">
      <HintPath>packages\ExcelNumberFormat.1.1.0\lib\net20\ExcelNumberFormat.dll</HintPath>
    </Reference>
    <Reference Include="Irony, Version=1.0.11.0, Culture=neutral, PublicKeyToken=ca48ace7223ead47, processorArchitecture=MSIL">
      <HintPath>packages\Irony.NetCore.1.0.11\lib\net461\Irony.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Bcl.AsyncInterfaces, Version=1.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Bcl.AsyncInterfaces.1.1.1\lib\net461\Microsoft.Bcl.AsyncInterfaces.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Bcl.HashCode, Version=1.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Bcl.HashCode.1.1.1\lib\net461\Microsoft.Bcl.HashCode.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.EntityFrameworkCore, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.EntityFrameworkCore.3.1.18\lib\netstandard2.0\Microsoft.EntityFrameworkCore.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.EntityFrameworkCore.Abstractions, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.EntityFrameworkCore.Abstractions.3.1.18\lib\netstandard2.0\Microsoft.EntityFrameworkCore.Abstractions.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.EntityFrameworkCore.Relational, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.EntityFrameworkCore.Relational.3.1.18\lib\netstandard2.0\Microsoft.EntityFrameworkCore.Relational.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Caching.Abstractions, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Caching.Abstractions.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Caching.Abstractions.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Caching.Memory, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Caching.Memory.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Caching.Memory.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Configuration, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Configuration.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Configuration.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Configuration.Abstractions, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Configuration.Abstractions.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Configuration.Abstractions.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Configuration.Binder, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Configuration.Binder.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Configuration.Binder.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.DependencyInjection, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.DependencyInjection.3.1.18\lib\net461\Microsoft.Extensions.DependencyInjection.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.DependencyInjection.Abstractions, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.DependencyInjection.Abstractions.3.1.18\lib\netstandard2.0\Microsoft.Extensions.DependencyInjection.Abstractions.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Logging, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Logging.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Logging.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Logging.Abstractions, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Logging.Abstractions.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Logging.Abstractions.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Options, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Options.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Options.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.Extensions.Primitives, Version=3.1.18.0, Culture=neutral, PublicKeyToken=adb9793829ddae60, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.Extensions.Primitives.3.1.18\lib\netstandard2.0\Microsoft.Extensions.Primitives.dll</HintPath>
    </Reference>
    <Reference Include="Microsoft.IO.RecyclableMemoryStream, Version=1.4.1.0, Culture=neutral, PublicKeyToken=31bf3856ad364e35, processorArchitecture=MSIL">
      <HintPath>packages\Microsoft.IO.RecyclableMemoryStream.1.4.1\lib\net46\Microsoft.IO.RecyclableMemoryStream.dll</HintPath>
    </Reference>
    <Reference Include="NLog, Version=5.0.0.0, Culture=neutral, PublicKeyToken=5120e14c03d0593c, processorArchitecture=MSIL">
      <HintPath>packages\NLog.5.2.5\lib\net46\NLog.dll</HintPath>
    </Reference>
    <Reference Include="Npgsql, Version=4.1.9.0, Culture=neutral, PublicKeyToken=5d8b90d52f46fda7, processorArchitecture=MSIL">
      <HintPath>packages\Npgsql.4.1.9\lib\net461\Npgsql.dll</HintPath>
    </Reference>
    <Reference Include="Npgsql.EntityFrameworkCore.PostgreSQL, Version=3.1.18.0, Culture=neutral, PublicKeyToken=5d8b90d52f46fda7, processorArchitecture=MSIL">
      <HintPath>packages\Npgsql.EntityFrameworkCore.PostgreSQL.3.1.18\lib\netstandard2.0\Npgsql.EntityFrameworkCore.PostgreSQL.dll</HintPath>
    </Reference>
    <Reference Include="PresentationCore" />
    <Reference Include="SixLabors.Fonts, Version=1.0.0.0, Culture=neutral, PublicKeyToken=d998eea7b14cab13, processorArchitecture=MSIL">
      <HintPath>packages\SixLabors.Fonts.1.0.0\lib\netstandard2.0\SixLabors.Fonts.dll</HintPath>
    </Reference>
    <Reference Include="System" />
    <Reference Include="System.Buffers, Version=4.0.3.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Buffers.4.5.1\lib\net461\System.Buffers.dll</HintPath>
    </Reference>
    <Reference Include="System.Collections.Immutable, Version=1.2.5.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Collections.Immutable.1.7.1\lib\net461\System.Collections.Immutable.dll</HintPath>
    </Reference>
    <Reference Include="System.ComponentModel.Annotations, Version=4.2.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.ComponentModel.Annotations.4.7.0\lib\net461\System.ComponentModel.Annotations.dll</HintPath>
    </Reference>
    <Reference Include="System.ComponentModel.Composition" />
    <Reference Include="System.ComponentModel.DataAnnotations" />
    <Reference Include="System.Configuration" />
    <Reference Include="System.Core" />
    <Reference Include="System.Diagnostics.DiagnosticSource, Version=4.0.5.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Diagnostics.DiagnosticSource.4.7.1\lib\net46\System.Diagnostics.DiagnosticSource.dll</HintPath>
    </Reference>
    <Reference Include="System.IO, Version=4.1.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.IO.4.3.0\lib\net462\System.IO.dll</HintPath>
      <Private>True</Private>
      <Private>True</Private>
    </Reference>
    <Reference Include="System.IO.Compression" />
    <Reference Include="System.IO.Packaging, Version=6.0.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.IO.Packaging.6.0.0\lib\net461\System.IO.Packaging.dll</HintPath>
    </Reference>
    <Reference Include="System.Memory, Version=4.0.1.1, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Memory.4.5.4\lib\net461\System.Memory.dll</HintPath>
    </Reference>
    <Reference Include="System.Net.Http, Version=4.1.1.3, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Net.Http.4.3.4\lib\net46\System.Net.Http.dll</HintPath>
      <Private>True</Private>
      <Private>True</Private>
    </Reference>
    <Reference Include="System.Numerics" />
    <Reference Include="System.Numerics.Vectors, Version=4.1.4.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Numerics.Vectors.4.5.0\lib\net46\System.Numerics.Vectors.dll</HintPath>
    </Reference>
    <Reference Include="System.Runtime, Version=4.1.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Runtime.4.3.0\lib\net462\System.Runtime.dll</HintPath>
      <Private>True</Private>
    </Reference>
    <Reference Include="System.Runtime.CompilerServices.Unsafe, Version=4.0.6.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Runtime.CompilerServices.Unsafe.4.7.1\lib\net461\System.Runtime.CompilerServices.Unsafe.dll</HintPath>
    </Reference>
    <Reference Include="System.Security" />
    <Reference Include="System.Security.Cryptography.Algorithms, Version=4.2.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Security.Cryptography.Algorithms.4.3.0\lib\net463\System.Security.Cryptography.Algorithms.dll</HintPath>
      <Private>True</Private>
    </Reference>
    <Reference Include="System.Security.Cryptography.Encoding, Version=4.0.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Security.Cryptography.Encoding.4.3.0\lib\net46\System.Security.Cryptography.Encoding.dll</HintPath>
      <Private>True</Private>
      <Private>True</Private>
    </Reference>
    <Reference Include="System.Security.Cryptography.Primitives, Version=4.0.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Security.Cryptography.Primitives.4.3.0\lib\net46\System.Security.Cryptography.Primitives.dll</HintPath>
      <Private>True</Private>
      <Private>True</Private>
    </Reference>
    <Reference Include="System.Security.Cryptography.X509Certificates, Version=4.1.1.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Security.Cryptography.X509Certificates.4.3.0\lib\net461\System.Security.Cryptography.X509Certificates.dll</HintPath>
      <Private>True</Private>
    </Reference>
    <Reference Include="System.Text.Encodings.Web, Version=4.0.4.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Text.Encodings.Web.4.6.0\lib\netstandard2.0\System.Text.Encodings.Web.dll</HintPath>
    </Reference>
    <Reference Include="System.Text.Json, Version=4.0.0.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Text.Json.4.6.0\lib\net461\System.Text.Json.dll</HintPath>
    </Reference>
    <Reference Include="System.Text.RegularExpressions, Version=4.1.0.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a, processorArchitecture=MSIL">
      <HintPath>packages\System.Text.RegularExpressions.4.3.1\lib\net463\System.Text.RegularExpressions.dll</HintPath>
      <Private>True</Private>
      <Private>True</Private>
    </Reference>
    <Reference Include="System.Threading.Tasks" />
    <Reference Include="System.Threading.Tasks.Extensions, Version=4.2.0.1, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.Threading.Tasks.Extensions.4.5.4\lib\net461\System.Threading.Tasks.Extensions.dll</HintPath>
    </Reference>
    <Reference Include="System.ValueTuple, Version=4.0.3.0, Culture=neutral, PublicKeyToken=cc7b13ffcd2ddd51, processorArchitecture=MSIL">
      <HintPath>packages\System.ValueTuple.4.5.0\lib\net47\System.ValueTuple.dll</HintPath>
    </Reference>
    <Reference Include="System.Web.Extensions" />
    <Reference Include="System.Xml.Linq" />
    <Reference Include="System.Data.DataSetExtensions" />
    <Reference Include="Microsoft.CSharp" />
    <Reference Include="System.Data" />
    <Reference Include="System.Deployment" />
    <Reference Include="System.Drawing" />
    <Reference Include="System.Windows.Forms" />
    <Reference Include="System.Xml" />
    <Reference Include="WindowsBase" />
    <Reference Include="XLParser, Version=1.5.2.0, Culture=neutral, PublicKeyToken=63397e1e46bb91b4, processorArchitecture=MSIL">
      <HintPath>packages\XLParser.1.5.2\lib\net461\XLParser.dll</HintPath>
    </Reference>
  </ItemGroup>
  <ItemGroup>
    <Compile Include="DBConnect.cs" />
    <Compile Include="FillExcel.cs" />
    <Compile Include="Form1.cs">
      <SubType>Form</SubType>
    </Compile>
    <Compile Include="Form1.Designer.cs">
      <DependentUpon>Form1.cs</DependentUpon>
    </Compile>
    <Compile Include="Form2.cs">
      <SubType>Form</SubType>
    </Compile>
    <Compile Include="Form2.Designer.cs">
      <DependentUpon>Form2.cs</DependentUpon>
    </Compile>
    <Compile Include="Form3.cs">
      <SubType>Form</SubType>
    </Compile>
    <Compile Include="Form3.Designer.cs">
      <DependentUpon>Form3.cs</DependentUpon>
    </Compile>
    <Compile Include="Form4.cs">
      <SubType>Form</SubType>
    </Compile>
    <Compile Include="Form4.Designer.cs">
      <DependentUpon>Form4.cs</DependentUpon>
    </Compile>
    <Compile Include="Form5.cs">
      <SubType>Form</SubType>
    </Compile>
    <Compile Include="Form5.Designer.cs">
      <DependentUpon>Form5.cs</DependentUpon>
    </Compile>
    <Compile Include="Form6.cs">
      <SubType>Form</SubType>
    </Compile>
    <Compile Include="Form6.Designer.cs">
      <DependentUpon>Form6.cs</DependentUpon>
    </Compile>
    <Compile Include="Form7.cs">
      <SubType>Form</SubType>
    </Compile>
    <Compile Include="Form7.Designer.cs">
      <DependentUpon>Form7.cs</DependentUpon>
    </Compile>
    <Compile Include="MethodReadLine.cs" />
    <Compile Include="MethodReadProd.cs" />
    <Compile Include="Program.cs" />
    <Compile Include="ProjectLogger.cs" />
    <Compile Include="Properties\AssemblyInfo.cs" />
    <Compile Include="Request.cs" />
    <Compile Include="SaAsDi.cs" />
    <Compile Include="TabledbClasses.cs" />
    <Compile Include="TextTemplate1.cs">
      <AutoGen>True</AutoGen>
      <DesignTime>True</DesignTime>
      <DependentUpon>TextTemplate1.tt</DependentUpon>
    </Compile>
    <EmbeddedResource Include="Form1.resx">
      <DependentUpon>Form1.cs</DependentUpon>
    </EmbeddedResource>
    <EmbeddedResource Include="Form2.resx">
      <DependentUpon>Form2.cs</DependentUpon>
    </EmbeddedResource>
    <EmbeddedResource Include="Form3.resx">
      <DependentUpon>Form3.cs</DependentUpon>
    </EmbeddedResource>
    <EmbeddedResource Include="Form4.resx">
      <DependentUpon>Form4.cs</DependentUpon>
    </EmbeddedResource>
    <EmbeddedResource Include="Form5.resx">
      <DependentUpon>Form5.cs</DependentUpon>
    </EmbeddedResource>
    <EmbeddedResource Include="Form6.resx">
      <DependentUpon>Form6.cs</DependentUpon>
    </EmbeddedResource>
    <EmbeddedResource Include="Form7.resx">
      <DependentUpon>Form7.cs</DependentUpon>
    </EmbeddedResource>
    <EmbeddedResource Include="Properties\Resources.resx">
      <Generator>ResXFileCodeGenerator</Generator>
      <LastGenOutput>Resources.Designer.cs</LastGenOutput>
      <SubType>Designer</SubType>
    </EmbeddedResource>
    <Compile Include="Properties\Resources.Designer.cs">
      <AutoGen>True</AutoGen>
      <DependentUpon>Resources.resx</DependentUpon>
      <DesignTime>True</DesignTime>
    </Compile>
    <None Include="packages.config" />
    <None Include="Properties\Settings.settings">
      <Generator>SettingsSingleFileGenerator</Generator>
      <LastGenOutput>Settings.Designer.cs</LastGenOutput>
    </None>
    <Compile Include="Properties\Settings.Designer.cs">
      <AutoGen>True</AutoGen>
      <DependentUpon>Settings.settings</DependentUpon>
      <DesignTimeSharedInput>True</DesignTimeSharedInput>
    </Compile>
  </ItemGroup>
  <ItemGroup>
    <None Include="App.config" />
  </ItemGroup>
  <ItemGroup>
    <Content Include="TextTemplate1.tt">
      <Generator>TextTemplatingFileGenerator</Generator>
      <LastGenOutput>TextTemplate1.cs</LastGenOutput>
    </Content>
  </ItemGroup>
  <ItemGroup>
    <Service Include="{508349B6-6B84-4DF5-91F0-309BEEBAD82D}" />
  </ItemGroup>
  <ItemGroup>
    <BootstrapperPackage Include=".NETFramework,Version=v4.7.2">
      <Visible>False</Visible>
      <ProductName>Microsoft .NET Framework 4.7.2 %28x86 и x64%29</ProductName>
      <Install>true</Install>
    </BootstrapperPackage>
    <BootstrapperPackage Include="Microsoft.Net.Framework.3.5.SP1">
      <Visible>False</Visible>
      <ProductName>.NET Framework 3.5 SP1</ProductName>
      <Install>false</Install>
    </BootstrapperPackage>
  </ItemGroup>
  <Import Project="$(MSBuildToolsPath)\Microsoft.CSharp.targets" />
</Project>