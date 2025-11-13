#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script to create a professional PowerPoint presentation about VPN Server Setup
on Windows Server and Linux
"""

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.dml.color import RGBColor

def create_vpn_presentation():
    """Create a comprehensive VPN presentation"""
    
    # Create presentation object
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Define color scheme
    TITLE_COLOR = RGBColor(0, 51, 102)  # Dark blue
    ACCENT_COLOR = RGBColor(0, 112, 192)  # Blue
    TEXT_COLOR = RGBColor(51, 51, 51)  # Dark gray
    
    # Slide 1: Title Slide
    slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout
    
    # Add background color
    background = slide1.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)
    
    # Title
    title_box = slide1.shapes.add_textbox(Inches(0.5), Inches(2), Inches(9), Inches(1.5))
    title_frame = title_box.text_frame
    title_frame.text = "Mise en place d'un serveur VPN sécurisé"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(44)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    title_para.alignment = PP_ALIGN.CENTER
    
    # Subtitle
    subtitle_box = slide1.shapes.add_textbox(Inches(0.5), Inches(3.5), Inches(9), Inches(1))
    subtitle_frame = subtitle_box.text_frame
    subtitle_frame.text = "sous Windows Server 2019 et Linux"
    subtitle_para = subtitle_frame.paragraphs[0]
    subtitle_para.font.size = Pt(32)
    subtitle_para.font.color.rgb = ACCENT_COLOR
    subtitle_para.alignment = PP_ALIGN.CENTER
    
    # Project info
    info_box = slide1.shapes.add_textbox(Inches(1), Inches(5), Inches(8), Inches(1))
    info_frame = info_box.text_frame
    info_frame.text = "Implémentation IPSec et SSL pour la connexion à distance des employés"
    info_para = info_frame.paragraphs[0]
    info_para.font.size = Pt(18)
    info_para.font.color.rgb = TEXT_COLOR
    info_para.alignment = PP_ALIGN.CENTER
    
    # Slide 2: Table of Contents
    slide2 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide2.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    # Title
    title_box = slide2.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "📋 Plan de la présentation"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    # Content
    content_box = slide2.shapes.add_textbox(Inches(1.5), Inches(1.8), Inches(7), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    toc_items = [
        "1. Contexte du projet",
        "2. Objectifs",
        "3. Outils et technologies",
        "4. Architecture réseau",
        "5. Configuration Windows Server 2019",
        "6. Configuration Linux (Ubuntu)",
        "7. Sécurisation et authentification",
        "8. Tests et validation",
        "9. Analyse du trafic",
        "10. Résultats et conclusion"
    ]
    
    for item in toc_items:
        p = content_frame.add_paragraph()
        p.text = item
        p.font.size = Pt(22)
        p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(12)
        p.level = 0
    
    # Slide 3: Contexte du projet
    slide3 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide3.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide3.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🧠 Contexte du projet"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide3.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    context_text = [
        "• Évolution du télétravail et des connexions à distance",
        "• Nécessité de sécuriser les communications entreprise",
        "• Protection des données sensibles en transit",
        "• Accès sécurisé aux ressources internes depuis l'extérieur",
        "",
        "🎯 Solution : VPN (Virtual Private Network)",
        "",
        "• Tunnel chiffré entre utilisateurs distants et réseau interne",
        "• Implémentation sur deux environnements :",
        "  - Windows Server 2019",
        "  - Linux (Ubuntu Server 22.04)"
    ]
    
    for i, text in enumerate(context_text):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        p.font.size = Pt(18)
        p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(8)
        if "🎯" in text:
            p.font.bold = True
            p.font.size = Pt(20)
            p.font.color.rgb = ACCENT_COLOR
    
    # Slide 4: Objectifs
    slide4 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide4.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide4.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "⚙️ Objectifs du projet"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide4.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    objectives = [
        "1️⃣ Comprendre le fonctionnement d'un VPN",
        "   • Protocoles IPSec et SSL/TLS",
        "   • Tunneling et encapsulation",
        "",
        "2️⃣ Installer et configurer un serveur VPN",
        "   • Windows Server 2019 (RRAS)",
        "   • Linux Ubuntu (OpenVPN / StrongSwan)",
        "",
        "3️⃣ Sécuriser les échanges",
        "   • Chiffrement fort (AES-256)",
        "   • Authentification multi-facteurs",
        "",
        "4️⃣ Tester la connexion distante",
        "",
        "5️⃣ Comparer les performances entre environnements"
    ]
    
    for i, text in enumerate(objectives):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(18)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 5: Outils et technologies
    slide5 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide5.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide5.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🧩 Outils et technologies utilisés"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    # Left column
    left_box = slide5.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(4), Inches(5))
    left_frame = left_box.text_frame
    left_frame.word_wrap = True
    
    left_content = [
        "🖥️ Virtualisation",
        "• VirtualBox / VMware",
        "",
        "🪟 Windows",
        "• Windows Server 2019",
        "• PowerShell",
        "• RRAS (Routing and Remote Access)",
        "",
        "🐧 Linux",
        "• Ubuntu Server 22.04",
        "• OpenVPN",
        "• StrongSwan (IPSec)"
    ]
    
    for i, text in enumerate(left_content):
        if i == 0:
            left_frame.text = text
            p = left_frame.paragraphs[0]
        else:
            p = left_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("🖥️", "🪟", "🐧")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(18)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Right column
    right_box = slide5.shapes.add_textbox(Inches(5.2), Inches(1.8), Inches(4), Inches(5))
    right_frame = right_box.text_frame
    right_frame.word_wrap = True
    
    right_content = [
        "🔍 Analyse et tests",
        "• Wireshark",
        "• Ping / Traceroute",
        "• Telnet / Netcat",
        "",
        "🔐 Sécurité",
        "• Certificats SSL/TLS",
        "• Clés IPSec",
        "• Authentification RADIUS",
        "",
        "📊 Documentation",
        "• Rapport technique",
        "• Schémas réseau",
        "• Procédures de configuration"
    ]
    
    for i, text in enumerate(right_content):
        if i == 0:
            right_frame.text = text
            p = right_frame.paragraphs[0]
        else:
            p = right_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("🔍", "🔐", "📊")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(18)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 6: Architecture réseau
    slide6 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide6.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide6.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🏗️ Architecture réseau"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide6.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    arch_content = [
        "📍 Topologie réseau sur VirtualBox/VMware",
        "",
        "1. Serveur Windows Server 2019",
        "   • IP: 192.168.1.10/24",
        "   • Rôle: Serveur VPN IPSec/L2TP",
        "",
        "2. Serveur Linux Ubuntu 22.04",
        "   • IP: 192.168.1.20/24",
        "   • Rôle: Serveur OpenVPN/StrongSwan",
        "",
        "3. Client distant (Windows/Linux)",
        "   • IP: DHCP ou 192.168.1.100/24",
        "   • Rôle: Client VPN",
        "",
        "4. Réseau interne",
        "   • Plage: 10.0.0.0/24",
        "   • Ressources: Serveurs de fichiers, applications"
    ]
    
    for i, text in enumerate(arch_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith("📍"):
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        elif text.startswith(("1.", "2.", "3.", "4.")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = TITLE_COLOR
        else:
            p.font.size = Pt(18)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 7: Configuration Windows Server
    slide7 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide7.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide7.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🪟 Configuration Windows Server 2019"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide7.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    win_content = [
        "📝 Étapes de configuration",
        "",
        "1. Installation du rôle Remote Access",
        "   • Server Manager → Add Roles → Remote Access",
        "   • Sélectionner DirectAccess and VPN (RAS)",
        "",
        "2. Configuration RRAS",
        "   • Routing and Remote Access → Configure",
        "   • Choisir 'Custom Configuration' → VPN Access",
        "",
        "3. Configuration des protocoles",
        "   • L2TP/IPSec avec clé pré-partagée",
        "   • PPTP (optionnel, moins sécurisé)",
        "   • IKEv2 (recommandé)",
        "",
        "4. Configuration du pool d'adresses",
        "   • Plage: 10.0.0.100 - 10.0.0.200",
        "",
        "5. Création des comptes utilisateurs VPN"
    ]
    
    for i, text in enumerate(win_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith("📝"):
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        elif text.startswith(("1.", "2.", "3.", "4.", "5.")):
            p.font.size = Pt(19)
            p.font.bold = True
            p.font.color.rgb = TITLE_COLOR
        else:
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(5)
    
    # Slide 8: PowerShell Commands
    slide8 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide8.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide8.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "💻 Commandes PowerShell"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    # Code box
    code_box = slide8.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    code_frame = code_box.text_frame
    code_frame.word_wrap = True
    
    ps_commands = [
        "# Installation du rôle Remote Access",
        "Install-WindowsFeature RemoteAccess -IncludeManagementTools",
        "Install-WindowsFeature DirectAccess-VPN -IncludeManagementTools",
        "Install-WindowsFeature Routing -IncludeManagementTools",
        "",
        "# Configuration VPN",
        "Install-RemoteAccess -VpnType Vpn",
        "",
        "# Configuration du pool d'adresses",
        "Set-VpnIPAddressAssignment -IPAddressRangeStart 10.0.0.100 `",
        "  -IPAddressRangeEnd 10.0.0.200",
        "",
        "# Création d'un utilisateur VPN",
        "New-LocalUser -Name 'vpnuser' -Password (ConvertTo-SecureString `",
        "  'P@ssw0rd!' -AsPlainText -Force)",
        "",
        "# Autoriser l'accès VPN",
        "Set-VpnUser -UserName 'vpnuser' -PassThru"
    ]
    
    for i, text in enumerate(ps_commands):
        if i == 0:
            code_frame.text = text
            p = code_frame.paragraphs[0]
        else:
            p = code_frame.add_paragraph()
            p.text = text
        
        p.font.name = 'Courier New'
        if text.startswith("#"):
            p.font.size = Pt(15)
            p.font.color.rgb = RGBColor(0, 128, 0)
            p.font.bold = True
        else:
            p.font.size = Pt(14)
            p.font.color.rgb = RGBColor(0, 0, 128)
        p.space_before = Pt(3)
    
    # Slide 9: Configuration Linux
    slide9 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide9.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide9.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🐧 Configuration Linux (Ubuntu)"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide9.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    linux_content = [
        "📝 Option 1 : OpenVPN (SSL/TLS)",
        "",
        "1. Installation",
        "   • sudo apt update && sudo apt install openvpn easy-rsa",
        "",
        "2. Configuration PKI (Public Key Infrastructure)",
        "   • Génération de l'autorité de certification (CA)",
        "   • Création des certificats serveur et clients",
        "",
        "3. Configuration du serveur",
        "   • Fichier: /etc/openvpn/server.conf",
        "   • Port: 1194 (UDP)",
        "   • Protocole: UDP/TCP",
        "",
        "📝 Option 2 : StrongSwan (IPSec)",
        "",
        "1. Installation",
        "   • sudo apt install strongswan strongswan-pki",
        "",
        "2. Configuration IPSec",
        "   • Fichier: /etc/ipsec.conf et /etc/ipsec.secrets"
    ]
    
    for i, text in enumerate(linux_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith("📝"):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        elif text.startswith(("1.", "2.", "3.")):
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = TITLE_COLOR
        else:
            p.font.size = Pt(16)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(4)
    
    # Slide 10: OpenVPN Commands
    slide10 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide10.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide10.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "💻 Commandes OpenVPN"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    code_box = slide10.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    code_frame = code_box.text_frame
    code_frame.word_wrap = True
    
    ovpn_commands = [
        "# Installation",
        "sudo apt update",
        "sudo apt install openvpn easy-rsa -y",
        "",
        "# Configuration Easy-RSA",
        "make-cadir ~/openvpn-ca",
        "cd ~/openvpn-ca",
        "./easyrsa init-pki",
        "./easyrsa build-ca nopass",
        "",
        "# Génération certificat serveur",
        "./easyrsa gen-req server nopass",
        "./easyrsa sign-req server server",
        "",
        "# Génération clé Diffie-Hellman",
        "./easyrsa gen-dh",
        "",
        "# Génération clé TLS",
        "openvpn --genkey --secret ta.key",
        "",
        "# Démarrage du service",
        "sudo systemctl start openvpn@server",
        "sudo systemctl enable openvpn@server"
    ]
    
    for i, text in enumerate(ovpn_commands):
        if i == 0:
            code_frame.text = text
            p = code_frame.paragraphs[0]
        else:
            p = code_frame.add_paragraph()
            p.text = text
        
        p.font.name = 'Courier New'
        if text.startswith("#"):
            p.font.size = Pt(15)
            p.font.color.rgb = RGBColor(0, 128, 0)
            p.font.bold = True
        else:
            p.font.size = Pt(13)
            p.font.color.rgb = RGBColor(0, 0, 128)
        p.space_before = Pt(3)
    
    # Slide 11: Sécurisation
    slide11 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide11.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide11.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🔒 Sécurisation et authentification"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    # Left column
    left_box = slide11.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(4.2), Inches(5))
    left_frame = left_box.text_frame
    left_frame.word_wrap = True
    
    sec_left = [
        "🔐 Chiffrement",
        "",
        "• AES-256-CBC/GCM",
        "• SHA-256/SHA-512 (HMAC)",
        "• RSA 2048/4096 bits",
        "• Perfect Forward Secrecy",
        "",
        "🔑 Authentification",
        "",
        "• Certificats X.509",
        "• Clés pré-partagées (PSK)",
        "• RADIUS/LDAP",
        "• Active Directory",
        "• Multi-facteurs (2FA)"
    ]
    
    for i, text in enumerate(sec_left):
        if i == 0:
            left_frame.text = text
            p = left_frame.paragraphs[0]
        else:
            p = left_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("🔐", "🔑")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Right column
    right_box = slide11.shapes.add_textbox(Inches(5.2), Inches(1.8), Inches(4), Inches(5))
    right_frame = right_box.text_frame
    right_frame.word_wrap = True
    
    sec_right = [
        "🛡️ Bonnes pratiques",
        "",
        "• Désactiver les protocoles faibles",
        "  (PPTP, MD5)",
        "",
        "• Utiliser des mots de passe forts",
        "",
        "• Renouveler les certificats",
        "  régulièrement",
        "",
        "• Activer les logs d'audit",
        "",
        "• Configurer le pare-feu",
        "",
        "• Limiter les tentatives",
        "  de connexion",
        "",
        "• Mettre à jour régulièrement"
    ]
    
    for i, text in enumerate(sec_right):
        if i == 0:
            right_frame.text = text
            p = right_frame.paragraphs[0]
        else:
            p = right_frame.add_paragraph()
            p.text = text
        
        if text.startswith("🛡️"):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 12: Tests et validation
    slide12 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide12.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide12.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "✅ Tests et validation"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide12.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    test_content = [
        "🧪 Tests de connectivité",
        "",
        "1. Connexion client VPN",
        "   • Windows: Paramètres → Réseau → VPN",
        "   • Linux: sudo openvpn --config client.ovpn",
        "",
        "2. Vérification de l'adresse IP",
        "   • ip addr show (Linux)",
        "   • ipconfig /all (Windows)",
        "",
        "3. Test de ping vers le réseau interne",
        "   • ping 10.0.0.1",
        "   • ping serveur-interne.local",
        "",
        "4. Test d'accès aux ressources",
        "   • Partages de fichiers (SMB/NFS)",
        "   • Applications web internes",
        "   • Bases de données",
        "",
        "5. Test de débit et latence",
        "   • iperf3 -c serveur-vpn",
        "   • speedtest-cli"
    ]
    
    for i, text in enumerate(test_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith("🧪"):
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        elif text.startswith(("1.", "2.", "3.", "4.", "5.")):
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = TITLE_COLOR
        else:
            p.font.size = Pt(16)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(5)
    
    # Slide 13: Analyse Wireshark
    slide13 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide13.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide13.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🔍 Analyse du trafic avec Wireshark"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide13.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    wireshark_content = [
        "📊 Objectifs de l'analyse",
        "",
        "• Vérifier le chiffrement des données",
        "• Identifier les protocoles utilisés",
        "• Détecter d'éventuelles failles de sécurité",
        "• Mesurer les performances",
        "",
        "🔎 Filtres Wireshark utiles",
        "",
        "• OpenVPN: udp.port == 1194 || tcp.port == 1194",
        "• IPSec: esp || isakmp || ipsec",
        "• L2TP: udp.port == 1701",
        "• IKEv2: udp.port == 500 || udp.port == 4500",
        "",
        "✅ Points de vérification",
        "",
        "• Handshake TLS/SSL visible",
        "• Données applicatives chiffrées (non lisibles)",
        "• Absence de fuites DNS",
        "• Intégrité des paquets (HMAC)"
    ]
    
    for i, text in enumerate(wireshark_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("📊", "🔎", "✅")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 14: Comparaison Windows vs Linux
    slide14 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide14.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide14.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "⚖️ Comparaison Windows vs Linux"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    # Left column - Windows
    left_box = slide14.shapes.add_textbox(Inches(0.5), Inches(1.8), Inches(4.5), Inches(5))
    left_frame = left_box.text_frame
    left_frame.word_wrap = True
    
    win_comp = [
        "🪟 Windows Server",
        "",
        "✅ Avantages:",
        "• Interface graphique intuitive",
        "• Intégration Active Directory",
        "• Support Microsoft officiel",
        "• Configuration simplifiée",
        "",
        "❌ Inconvénients:",
        "• Coût de licence élevé",
        "• Ressources système importantes",
        "• Moins flexible",
        "• Mises à jour fréquentes"
    ]
    
    for i, text in enumerate(win_comp):
        if i == 0:
            left_frame.text = text
            p = left_frame.paragraphs[0]
        else:
            p = left_frame.add_paragraph()
            p.text = text
        
        if text.startswith("🪟"):
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = TITLE_COLOR
        elif text.startswith(("✅", "❌")):
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(16)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Right column - Linux
    right_box = slide14.shapes.add_textbox(Inches(5), Inches(1.8), Inches(4.5), Inches(5))
    right_frame = right_box.text_frame
    right_frame.word_wrap = True
    
    linux_comp = [
        "🐧 Linux Ubuntu",
        "",
        "✅ Avantages:",
        "• Gratuit et open-source",
        "• Léger et performant",
        "• Très flexible et personnalisable",
        "• Communauté active",
        "",
        "❌ Inconvénients:",
        "• Courbe d'apprentissage",
        "• Configuration en ligne de commande",
        "• Support communautaire",
        "• Documentation dispersée"
    ]
    
    for i, text in enumerate(linux_comp):
        if i == 0:
            right_frame.text = text
            p = right_frame.paragraphs[0]
        else:
            p = right_frame.add_paragraph()
            p.text = text
        
        if text.startswith("🐧"):
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = TITLE_COLOR
        elif text.startswith(("✅", "❌")):
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(16)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 15: Résultats attendus
    slide15 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide15.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide15.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🎯 Résultats attendus"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide15.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    results_content = [
        "✅ Connexion distante sécurisée",
        "   • Tunnel VPN opérationnel entre client et serveur",
        "   • Authentification réussie des utilisateurs",
        "   • Stabilité de la connexion",
        "",
        "🔐 Chiffrement des données",
        "   • Trafic entièrement chiffré (AES-256)",
        "   • Vérification via Wireshark",
        "   • Aucune fuite de données en clair",
        "",
        "🌐 Accès aux ressources internes",
        "   • Accès complet au réseau interne (10.0.0.0/24)",
        "   • Partages de fichiers accessibles",
        "   • Applications web internes fonctionnelles",
        "",
        "📊 Performance acceptable",
        "   • Latence < 50ms",
        "   • Débit > 50 Mbps",
        "   • Pas de perte de paquets significative"
    ]
    
    for i, text in enumerate(results_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("✅", "🔐", "🌐", "📊")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 16: Problèmes courants
    slide16 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide16.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide16.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "⚠️ Problèmes courants et solutions"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide16.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    problems_content = [
        "❌ Problème: Connexion refusée",
        "✅ Solution: Vérifier pare-feu et ports ouverts",
        "   • Windows: netsh advfirewall firewall add rule...",
        "   • Linux: sudo ufw allow 1194/udp",
        "",
        "❌ Problème: Authentification échouée",
        "✅ Solution: Vérifier certificats et identifiants",
        "   • Dates de validité des certificats",
        "   • Permissions des fichiers de clés",
        "",
        "❌ Problème: Pas d'accès au réseau interne",
        "✅ Solution: Configurer le routage et NAT",
        "   • Windows: Enable-NetNatTranslation",
        "   • Linux: iptables -t nat -A POSTROUTING...",
        "",
        "❌ Problème: Performances dégradées",
        "✅ Solution: Optimiser la configuration",
        "   • Changer de protocole (UDP vs TCP)",
        "   • Ajuster la MTU",
        "   • Utiliser la compression"
    ]
    
    for i, text in enumerate(problems_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith("❌"):
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = RGBColor(220, 20, 60)
        elif text.startswith("✅"):
            p.font.size = Pt(18)
            p.font.bold = True
            p.font.color.rgb = RGBColor(34, 139, 34)
        else:
            p.font.size = Pt(15)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(5)
    
    # Slide 17: Documentation
    slide17 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide17.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide17.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "📚 Documentation et livrables"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide17.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    doc_content = [
        "📄 Rapport technique",
        "• Introduction et contexte",
        "• Architecture détaillée",
        "• Procédures de configuration pas à pas",
        "• Captures d'écran",
        "• Résultats des tests",
        "• Analyse comparative",
        "",
        "🗺️ Schémas réseau",
        "• Topologie complète",
        "• Flux de données",
        "• Plan d'adressage IP",
        "",
        "📋 Guides d'utilisation",
        "• Guide administrateur",
        "• Guide utilisateur client VPN",
        "• Procédures de dépannage",
        "",
        "💾 Fichiers de configuration",
        "• Scripts PowerShell",
        "• Fichiers OpenVPN (.ovpn)",
        "• Configurations StrongSwan"
    ]
    
    for i, text in enumerate(doc_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("📄", "🗺️", "📋", "💾")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 18: Améliorations futures
    slide18 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide18.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(255, 255, 255)
    
    title_box = slide18.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🚀 Améliorations futures"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide18.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    future_content = [
        "🔐 Sécurité avancée",
        "• Authentification multi-facteurs (MFA)",
        "• Intégration avec Azure AD / Google Workspace",
        "• Certificats avec renouvellement automatique",
        "• IDS/IPS pour détecter les intrusions",
        "",
        "📊 Monitoring et supervision",
        "• Dashboard de monitoring (Grafana)",
        "• Alertes en temps réel",
        "• Logs centralisés (ELK Stack)",
        "• Métriques de performance",
        "",
        "⚡ Haute disponibilité",
        "• Cluster de serveurs VPN",
        "• Load balancing",
        "• Failover automatique",
        "",
        "🌍 Déploiement multi-sites",
        "• VPN site-to-site",
        "• Interconnexion de plusieurs bureaux",
        "• Optimisation WAN"
    ]
    
    for i, text in enumerate(future_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("🔐", "📊", "⚡", "🌍")):
            p.font.size = Pt(20)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 19: Conclusion
    slide19 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide19.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(240, 248, 255)
    
    title_box = slide19.shapes.add_textbox(Inches(0.5), Inches(0.5), Inches(9), Inches(0.8))
    title_frame = title_box.text_frame
    title_frame.text = "🎓 Conclusion"
    title_para = title_frame.paragraphs[0]
    title_para.font.size = Pt(36)
    title_para.font.bold = True
    title_para.font.color.rgb = TITLE_COLOR
    
    content_box = slide19.shapes.add_textbox(Inches(0.8), Inches(1.8), Inches(8.4), Inches(4.5))
    content_frame = content_box.text_frame
    content_frame.word_wrap = True
    
    conclusion_content = [
        "✅ Objectifs atteints",
        "",
        "• Mise en place réussie d'un VPN sécurisé sur deux plateformes",
        "• Compréhension approfondie des protocoles VPN",
        "• Maîtrise des outils de configuration et d'analyse",
        "• Sécurisation efficace des communications distantes",
        "",
        "💡 Compétences acquises",
        "",
        "• Administration Windows Server et Linux",
        "• Configuration réseau avancée",
        "• Cryptographie et sécurité",
        "• Analyse de trafic réseau",
        "• Résolution de problèmes",
        "",
        "🌟 Impact pour l'entreprise",
        "",
        "• Télétravail sécurisé pour les employés",
        "• Protection des données sensibles",
        "• Flexibilité et productivité accrues"
    ]
    
    for i, text in enumerate(conclusion_content):
        if i == 0:
            content_frame.text = text
            p = content_frame.paragraphs[0]
        else:
            p = content_frame.add_paragraph()
            p.text = text
        
        if text.startswith(("✅", "💡", "🌟")):
            p.font.size = Pt(22)
            p.font.bold = True
            p.font.color.rgb = ACCENT_COLOR
        else:
            p.font.size = Pt(17)
            p.font.color.rgb = TEXT_COLOR
        p.space_before = Pt(6)
    
    # Slide 20: Questions
    slide20 = prs.slides.add_slide(prs.slide_layouts[6])
    background = slide20.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 51, 102)
    
    # Questions text
    questions_box = slide20.shapes.add_textbox(Inches(1), Inches(2.5), Inches(8), Inches(2))
    questions_frame = questions_box.text_frame
    questions_frame.text = "❓ Questions ?"
    questions_para = questions_frame.paragraphs[0]
    questions_para.font.size = Pt(72)
    questions_para.font.bold = True
    questions_para.font.color.rgb = RGBColor(255, 255, 255)
    questions_para.alignment = PP_ALIGN.CENTER
    
    # Thank you text
    thanks_box = slide20.shapes.add_textbox(Inches(1), Inches(5), Inches(8), Inches(1))
    thanks_frame = thanks_box.text_frame
    thanks_frame.text = "Merci de votre attention !"
    thanks_para = thanks_frame.paragraphs[0]
    thanks_para.font.size = Pt(36)
    thanks_para.font.color.rgb = RGBColor(255, 255, 255)
    thanks_para.alignment = PP_ALIGN.CENTER
    
    # Save presentation
    output_file = '/vercel/sandbox/VPN_Server_Presentation.pptx'
    prs.save(output_file)
    print(f"✅ Présentation créée avec succès: {output_file}")
    print(f"📊 Nombre de diapositives: {len(prs.slides)}")
    return output_file

if __name__ == "__main__":
    create_vpn_presentation()
