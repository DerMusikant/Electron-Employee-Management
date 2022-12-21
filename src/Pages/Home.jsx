import { Left } from '../Components/Left/Left.jsx'
import { Main } from '../Components/Main/Main.jsx'

export const Home = () =>
    <>
        <div class="pane-group">
            <div class="pane-sm sidebar">
                <Left />
            </div>
            <div class="pane">
                <Main />
            </div>
        </div>
    </>