#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
신성통상 AWS KMS 구성 현황 및 권장사항 PPT 생성
ARN: arn:aws:kms:ap-northeast-2:804812023181:key/ab360511-0205-459c-a360-ff2a1b95b842
"""
import sys
import os

PROJECT_ROOT = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, PROJECT_ROOT)

from server import create_presentation

result = create_presentation(
    cover_title="신성통상 AWS KMS\n구성 현황 및 권장사항",
    cover_subtitle="암호화 키 관리 정책 | DB 마이그레이션 프로젝트",
    sections=[
        {
            "section_title": "1. KMS 구성 현황",
            "slides": [
                {
                    "layout": "numbered_list",
                    "title": "1-1. KMS 키 구성 현황",
                    "description": "AWS KMS | ap-northeast-2",
                    "body_title": "신성통상 Customer Managed Key 현황",
                    "content": {
                        "items": [
                            {
                                "title": "키 ARN",
                                "desc": "arn:aws:kms:ap-northeast-2:804812023181:key/ab360511-0205-459c-a360-ff2a1b95b842"
                            },
                            {
                                "title": "리전 / 계정",
                                "desc": "ap-northeast-2 (서울) / 계정 ID: 804812023181"
                            },
                            {
                                "title": "키 유형",
                                "desc": "대칭키 (Symmetric) — 암호화 및 복호화 전용 / 상태: Enabled"
                            },
                            {
                                "title": "암호화 알고리즘",
                                "desc": "SYMMETRIC_DEFAULT (AES-256-GCM) — FIPS 140-2 레벨 3 인증 HSM 보관"
                            },
                            {
                                "title": "주요 사용처",
                                "desc": "AWS Glue ETL 작업 데이터 암호화 (메인), Amazon Aurora, AWS DMS"
                            }
                        ]
                    }
                },
                {
                    "layout": "3_cards",
                    "title": "1-2. KMS 연동 서비스",
                    "description": "DB 마이그레이션 암호화 적용 서비스",
                    "body_title": "암호화 적용 AWS 서비스",
                    "content": {
                        "card_1": {
                            "title": "AWS Glue",
                            "body": "주요 사용 서비스\nETL 작업 데이터 소스/대상 암호화\nGlue 카탈로그 메타데이터 보호\n보안 구성에 CMK 등록",
                            "search_q": "AWS Glue data integration"
                        },
                        "card_2": {
                            "title": "Amazon Aurora",
                            "body": "DB 인스턴스 저장 데이터 암호화\n스냅샷 및 백업 자동 암호화\n동일 CMK 재사용 가능",
                            "search_q": "Amazon Aurora PostgreSQL database"
                        },
                        "card_3": {
                            "title": "AWS DMS",
                            "body": "복제 인스턴스 스토리지 암호화\n마이그레이션 데이터 보호\nCMK 기반 암호화 적용",
                            "search_q": "AWS DMS database migration service"
                        }
                    }
                }
            ]
        },
        {
            "section_title": "2. 암호화 방식",
            "slides": [
                {
                    "layout": "challenge_solution",
                    "title": "2-1. 암호화 알고리즘 및 방식",
                    "description": "AES-256-GCM / Envelope Encryption",
                    "body_title": "AWS KMS 암호화 기술 구조",
                    "content": {
                        "challenge": {
                            "title": "알고리즘 (SYMMETRIC_DEFAULT)",
                            "body": "AES-256-GCM\n\n• 대칭키 암호화 방식\n• 256비트 키 길이\n• GCM 모드 (인증 암호화)\n• FIPS 140-2 레벨 3 인증\n• CMK는 AWS KMS HSM에 안전 보관"
                        },
                        "solution": {
                            "title": "Envelope Encryption",
                            "body": "2중 키 암호화 구조\n\n• 데이터 → DEK(데이터 암호화 키)로 암호화\n• DEK → CMK(마스터 키)로 암호화\n• 암호화된 DEK만 저장소에 보관\n• CMK는 항상 KMS 내부에만 존재\n• 대용량 데이터도 효율적 처리 가능"
                        }
                    }
                },
                {
                    "layout": "process_arrow",
                    "title": "2-2. Envelope Encryption 동작 흐름",
                    "description": "GenerateDataKey → Encrypt → Store → Decrypt",
                    "body_title": "데이터 암호화 / 복호화 프로세스",
                    "content": {
                        "steps": [
                            {
                                "title": "DEK 발급",
                                "body": "KMS GenerateDataKey 호출\n평문 DEK + 암호화된 DEK 반환"
                            },
                            {
                                "title": "데이터 암호화",
                                "body": "평문 DEK로 실제 데이터를 로컬에서 AES-256 암호화"
                            },
                            {
                                "title": "평문 DEK 삭제",
                                "body": "평문 DEK를 메모리에서 즉시 삭제\n보안 위협 원천 제거"
                            },
                            {
                                "title": "저장",
                                "body": "암호화된 데이터 + 암호화된 DEK 함께 저장"
                            },
                            {
                                "title": "복호화",
                                "body": "KMS Decrypt로 DEK 복호화\n→ DEK로 데이터 복호화"
                            }
                        ]
                    }
                }
            ]
        },
        {
            "section_title": "3. 자동 키 교체 정책",
            "slides": [
                {
                    "layout": "timeline_steps",
                    "title": "3-1. 자동 키 교체 주기 및 동작",
                    "description": "365일 / 연 1회 자동 교체",
                    "body_title": "AWS KMS 자동 키 교체 (Automatic Key Rotation)",
                    "content": {
                        "steps": [
                            {
                                "date": "교체 활성화",
                                "desc": "AWS 콘솔 또는 CLI에서\n자동 키 교체 Enable\n최초 1년 후부터 시작"
                            },
                            {
                                "date": "D+365일",
                                "desc": "새 키 재료(Key Material)\n자동 생성\n키 ID·ARN 동일 유지"
                            },
                            {
                                "date": "이전 키 보존",
                                "desc": "이전 키 재료 삭제 없음\n기존 암호화 데이터\n복호화 계속 가능"
                            },
                            {
                                "date": "신규 암호화",
                                "desc": "교체 이후 신규 암호화는\n최신 키 재료 자동 사용\n코드 변경 불필요"
                            },
                            {
                                "date": "반복",
                                "desc": "매 365일마다\n동일 과정 자동 반복\n무중단 운영"
                            }
                        ]
                    }
                },
                {
                    "layout": "grid_2x2",
                    "title": "3-2. 키 교체 핵심 원리",
                    "description": "Key Material Rotation",
                    "body_title": "자동 키 교체 상세 동작 원리",
                    "content": {
                        "item1": {
                            "title": "키 ID 불변",
                            "body": "교체 후에도 ARN / 키 ID 동일 유지\n→ 애플리케이션 코드 변경 없이 무중단 적용"
                        },
                        "item2": {
                            "title": "키 재료만 교체",
                            "body": "내부 Key Material만 새로 생성\n→ 보안 강화 + 장기 노출 위험 감소"
                        },
                        "item3": {
                            "title": "이전 키 재료 보존",
                            "body": "이전 Key Material 자동 보관\n→ 교체 전 암호화 데이터 영구 복호화 가능"
                        },
                        "item4": {
                            "title": "재암호화 없음",
                            "body": "기존 데이터 자동 재암호화 없음\n→ 필요 시 ReEncrypt API 수동 호출"
                        }
                    }
                }
            ]
        },
        {
            "section_title": "4. 권장사항",
            "slides": [
                {
                    "layout": "checklist_2col",
                    "title": "4-1. KMS 운영 권장사항",
                    "description": "Best Practices",
                    "body_title": "신성통상 KMS 권장 설정 및 운영 정책",
                    "content": {
                        "summary": "0/5 Passed    5 Warning",
                        "items": [
                            {
                                "title": "자동 키 교체 활성화",
                                "status": "todo",
                                "subitems": [
                                    {"text": "AWS 콘솔 → KMS → 키(ab360511) → 키 교체 탭 확인", "badge": "HIGH"},
                                    {"text": "자동 교체 Enable 후 다음 교체 예정일 확인"}
                                ]
                            },
                            {
                                "title": "CloudTrail 모니터링 설정",
                                "status": "todo",
                                "subitems": [
                                    {"text": "KMS API 호출 로그 CloudTrail 활성화"},
                                    {"text": "Decrypt / GenerateDataKey 이상 호출 알림 설정", "badge": "MEDIUM"}
                                ]
                            },
                            {
                                "title": "최소 권한 IAM 정책 적용",
                                "status": "todo",
                                "subitems": [
                                    {"text": "서비스별 kms:GenerateDataKey / kms:Decrypt 권한 분리", "badge": "HIGH"},
                                    {"text": "Glue / Aurora / DMS 각각 별도 정책 적용"}
                                ]
                            },
                            {
                                "title": "키 삭제 대기 기간 설정",
                                "status": "todo",
                                "subitems": [
                                    {"text": "불필요 키 삭제 시 최소 7일 (권장 30일) 대기"},
                                    {"text": "삭제 전 암호화 데이터 복호화 여부 확인"}
                                ]
                            },
                            {
                                "title": "Multi-Region 키 복제 검토",
                                "status": "todo",
                                "subitems": [
                                    {"text": "DR 리전(ap-northeast-1) 대비 복제 필요성 평가"},
                                    {"text": "Multi-Region 키 적용 시 비용 및 운영 방안 검토"}
                                ]
                            }
                        ]
                    }
                }
            ]
        }
    ],
    output_name="ssts_kms_policy"
)

print(f"생성 결과: {result}")
