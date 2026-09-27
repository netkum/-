# ==============================================================
# build.ps1 - 팀 스케줄러 엑셀 파일 자동 생성
#   이 파일이 있는 폴더(예: F:\claude\Scheduler)를 기준으로 상대 경로 사용
#   생성: Template\업무일정_Template.xlsm, Master\연구팀_Master.xlsm
#   필요: Excel 설치 + [VBA 프로젝트 개체 모델에 안전하게 액세스] 허용
# ==============================================================
$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$vba  = Join-Path $root 'vba'
foreach ($d in 'Master', 'Member', 'Template') {
    New-Item -ItemType Directory -Force -Path (Join-Path $root $d) | Out-Null
}

function Set-Code($component, $path) {
    $cm = $component.CodeModule
    if ($cm.CountOfLines -gt 0) { $cm.DeleteLines(1, $cm.CountOfLines) }
    $text = [System.IO.File]::ReadAllText($path, [System.Text.Encoding]::UTF8)
    $cm.AddFromString($text)
}

function Get-WorkbookModule($wb) {
    $name = $wb.CodeName
    if ($name) { return $wb.VBProject.VBComponents.Item($name) }
    foreach ($c in $wb.VBProject.VBComponents) {
        if ($c.Type -eq 100) {
            try { $null = $c.Properties.Item('IsAddin'); return $c } catch { }
        }
    }
    throw 'ThisWorkbook 모듈을 찾지 못했습니다.'
}

function New-Book($excel, $pkg, [string[]]$forms, $setupMacro, $savePath) {
    $wb = $excel.Workbooks.Add()
    try {
        $vbp = $wb.VBProject
        $null = $vbp.VBComponents.Count
    } catch {
        $wb.Close($false)
        throw "VBA 프로젝트 접근이 막혀 있습니다.`n엑셀 > 파일 > 옵션 > 보안 센터 > 보안 센터 설정 > 매크로 설정 >`n'VBA 프로젝트 개체 모델에 안전하게 액세스할 수 있음'을 체크한 뒤 다시 실행하세요.`n(회사 정책으로 막혀 있으면 README의 수동 설치 방법을 사용하세요)"
    }
    $dir = Join-Path $vba $pkg
    # 1) 사용자 정의 폼 (MSForms 참조가 먼저 추가되도록 가장 먼저)
    foreach ($f in $forms) {
        $c = $vbp.VBComponents.Add(3)
        $c.Name = $f
        Set-Code $c (Join-Path $dir "$f.txt")
    }
    # 2) 클래스 / 모듈 (CP949 파일)
    Get-ChildItem $dir -Filter *.cls | ForEach-Object { $null = $vbp.VBComponents.Import($_.FullName) }
    Get-ChildItem $dir -Filter *.bas | ForEach-Object { $null = $vbp.VBComponents.Import($_.FullName) }
    # 3) ThisWorkbook
    Set-Code (Get-WorkbookModule $wb) (Join-Path $dir 'ThisWorkbook.txt')
    # 4) 시트·표·버튼 구성
    $excel.Run("'" + $wb.Name + "'!" + $setupMacro)
    # 5) 저장 (52 = 매크로 사용 통합 문서)
    if (Test-Path $savePath) { Remove-Item $savePath -Force }
    $wb.SaveAs($savePath, 52)
    $wb.Close($false)
    Write-Host "  만들었습니다: $savePath"
}

Write-Host '팀 스케줄러 파일을 만듭니다...' -ForegroundColor Cyan
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $true
$excel.DisplayAlerts = $false
try {
    New-Book $excel 'member' @('frmTask', 'frmRequest', 'frmPhase') 'SetupTemplateSilent' (Join-Path $root 'Template\업무일정_Template.xlsm')
    New-Book $excel 'master' @() 'SetupMasterSilent' (Join-Path $root 'Master\연구팀_Master.xlsm')
    Write-Host ''
    Write-Host '완료! 다음 순서로 사용하세요.' -ForegroundColor Green
    Write-Host '  1) Template\업무일정_Template.xlsm 을 Member 폴더에 복사 → 이름_업무일정.xlsm 으로 이름 변경 (팀원마다)'
    Write-Host '  2) Master\연구팀_Master.xlsm 열기 → ProjectMaster에 프로젝트 등록 → [ProjectList 배포]'
    Write-Host '  3) 팀원 파일을 열어 Task 등록·저장 → Master에서 [전체 동기화]'
} catch {
    Write-Host ''
    Write-Host ('오류: ' + $_.Exception.Message) -ForegroundColor Red
} finally {
    $excel.DisplayAlerts = $true
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
}
