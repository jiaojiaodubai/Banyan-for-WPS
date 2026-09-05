// 官方 checkbox 交互行为及颜色（右侧为当前方案实现情况）：
// 1. 默认边框/未选中：#c1c1c1         （原生，已实现）
// 2. 选中主色：#1e5fc7                （accent-color，已实现）
// 3. hover 填充（已勾选）：#458bfa     （accent-color，已实现）
// 4. hover 填充（未勾选）：#ececec     （原生控件难以实现，未实现）
// 5. 勾选符号粗细/比例/动画           （由系统/浏览器决定，部分可控）
// 6. 尺寸随字号缩放                   （em 单位，已实现）
// 7. label 垂直居中                   （flex + line-height，已实现）
const checkboxTemplate = document.createElement("template")
checkboxTemplate.innerHTML = `
  <style>
    :host {
      display: block;
      font-size: inherit;
      line-height: 1.2;
    }

    .banyan-checkbox {
      display: inline-flex;
      align-items: center;
      gap: 0.14em;
      cursor: pointer;
      user-select: none;
    }

    .banyan-checkbox__label {
      display: inline-block;
      line-height: 1.2;
    }

    .banyan-checkbox__input {
      display: block;
      width: 1em;
      height: 1em;
      margin: 0;
      cursor: pointer;
      accent-color: #1e5fc7;
    }

    .banyan-checkbox__input:hover {
      accent-color: #458bfa;
    }

    .banyan-checkbox__input:focus-visible {
      outline: 2px solid #458bfa;
      outline-offset: 1px;
    }

    .banyan-checkbox__input:disabled {
      cursor: not-allowed;
      opacity: 0.6;
    }
  </style>
  <label class="banyan-checkbox">
    <input class="banyan-checkbox__input" type="checkbox" />
    <span class="banyan-checkbox__label"></span>
  </label>
`

export class BanyanCheckboxElement extends HTMLElement {
  private labelElement: HTMLSpanElement
  private inputElement: HTMLInputElement

  constructor() {
    super()
    const shadow = this.attachShadow({ mode: "open" })
    shadow.appendChild(checkboxTemplate.content.cloneNode(true))
    this.inputElement = shadow.querySelector(".banyan-checkbox__input") as HTMLInputElement
    this.labelElement = shadow.querySelector(".banyan-checkbox__label") as HTMLSpanElement
  }

  static get observedAttributes() {
    return ["checked", "disabled", "label"]
  }

  connectedCallback() {
    this.syncFromAttributes()
    this.inputElement.addEventListener("change", this.onInputChange)
  }

  disconnectedCallback() {
    this.inputElement.removeEventListener("change", this.onInputChange)
  }

  attributeChangedCallback() {
    this.syncFromAttributes()
  }

  get checked(): boolean {
    return this.inputElement.checked
  }

  set checked(value: boolean) {
    const nextChecked = Boolean(value)
    this.inputElement.checked = nextChecked
    this.toggleAttribute("checked", nextChecked)
  }

  get disabled(): boolean {
    return this.inputElement.disabled
  }

  set disabled(value: boolean) {
    const nextDisabled = Boolean(value)
    this.inputElement.disabled = nextDisabled
    this.toggleAttribute("disabled", nextDisabled)
  }

  get label(): string {
    return this.labelElement.textContent || ""
  }

  set label(value: string) {
    this.labelElement.textContent = value
    this.setAttribute("label", value)
  }

  focus(options?: FocusOptions) {
    this.inputElement.focus(options)
  }

  private onInputChange = () => {
    this.toggleAttribute("checked", this.inputElement.checked)
    this.dispatchEvent(new Event("change", { bubbles: true, composed: true }))
  }

  private syncFromAttributes() {
    this.inputElement.checked = this.hasAttribute("checked")
    this.inputElement.disabled = this.hasAttribute("disabled")
    this.labelElement.textContent = this.getAttribute("label") || ""
  }
}

if (!customElements.get("banyan-checkbox")) {
  customElements.define("banyan-checkbox", BanyanCheckboxElement)
}
